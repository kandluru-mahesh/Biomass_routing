import streamlit as st
import pandas as pd
from geopy.distance import geodesic
from ortools.constraint_solver import pywrapcp, routing_enums_pb2
from io import BytesIO

st.set_page_config(page_title="Biomass Route Optimizer", layout="wide")
st.title("🚜 Biomass Route Optimizer")

# --- 🔄 Split Function for Oversized Loads ---
def split_oversized_suppliers(df, capacity):
    new_rows = []
    for idx, row in df.iterrows():
        qty = row['Biomass Quantity (kg)']
        if qty <= capacity:
            new_rows.append(row)
        else:
            n_chunks = qty // capacity
            remainder = qty % capacity

            # Full-capacity chunks
            for i in range(int(n_chunks)):
                chunk = row.copy()
                chunk['Biomass Quantity (kg)'] = capacity
                chunk['Supplier Name (Farmer Name)'] += f" (Split-{i+1})"
                new_rows.append(chunk)

            # Remaining load
            if remainder > 0:
                chunk = row.copy()
                chunk['Biomass Quantity (kg)'] = remainder
                chunk['Supplier Name (Farmer Name)'] += f" (Split-R)"
                new_rows.append(chunk)
    return pd.DataFrame(new_rows)

# --- 📍 Warehouse & Capacity Input ---
with st.expander("📍 Enter Warehouse Location"):
    warehouse_lat = st.number_input("Warehouse Latitude", format="%.6f", value=14.083970)
    warehouse_lon = st.number_input("Warehouse Longitude", format="%.6f", value=79.794420)
warehouse = ('Warehouse', warehouse_lat, warehouse_lon)

tractor_capacity = st.number_input("🛻 Tractor Capacity (kg):", min_value=100, step=100, value=2000)

# --- 📤 Upload Supplier File ---
st.subheader("📤 Upload Supplier Data (.csv or .xlsx)")
uploaded_file = st.file_uploader("Choose a file", type=['xlsx', 'csv'])
st.caption("💡 Required columns: 'Supplier Name (Farmer Name)', 'Latitude of the location', 'Longitude of the location', 'Biomass Type', 'Biomass Quantity'")

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df_full = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith('.xlsx'):
            try:
                import openpyxl
                df_full = pd.read_excel(uploaded_file, engine='openpyxl')
            except ImportError:
                st.error("❌ You need to install the `openpyxl` package to read Excel files.\n\nRun: `pip install openpyxl`")
                st.stop()
        else:
            st.error("❌ Unsupported file format. Please upload a .csv or .xlsx file.")
            st.stop()

        # ✅ Validate required columns
        required_cols = {
            'Supplier Name (Farmer Name)',
            'Latitude of the location',
            'Longitude of the location',
            'Biomass Type',
            'Biomass Quantity (kg)'
        }

        if not required_cols.issubset(df_full.columns):
            st.error("❌ Your file must contain the following columns exactly:\n\n"
                     "- Supplier Name (Farmer Name)\n"
                     "- Latitude of the location\n"
                     "- Longitude of the location\n"
                     "- Biomass Type\n"
                     "- Biomass Quantity (kg)")
            st.stop()

        # 🌿 Select Biomass Type
        biomass_type = st.selectbox("🌿 Select Biomass Type:", df_full['Biomass Type'].unique())

        if st.button("Generate Routes"):
            df = df_full[df_full['Biomass Type'] == biomass_type].reset_index(drop=True)
            df = split_oversized_suppliers(df, tractor_capacity)

            # 📍 Build distance matrix
            locations = [(warehouse[1], warehouse[2])] + list(zip(df['Latitude of the location'], df['Longitude of the location']))
            demands = [0] + df['Biomass Quantity'].tolist()

            def compute_distance_matrix(locations):
                matrix = []
                for loc1 in locations:
                    row = []
                    for loc2 in locations:
                        dist = geodesic(loc1, loc2).km
                        row.append(int(dist * 1000))  # meters
                    matrix.append(row)
                return matrix

            distance_matrix = compute_distance_matrix(locations)
            num_vehicles = len(df)
            depot = 0
            manager = pywrapcp.RoutingIndexManager(len(distance_matrix), num_vehicles, depot)
            routing = pywrapcp.RoutingModel(manager)

            def distance_callback(from_index, to_index):
                return distance_matrix[manager.IndexToNode(from_index)][manager.IndexToNode(to_index)]
            transit_callback_index = routing.RegisterTransitCallback(distance_callback)
            routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)

            def demand_callback(from_index):
                return demands[manager.IndexToNode(from_index)]
            demand_callback_index = routing.RegisterUnaryTransitCallback(demand_callback)
            routing.AddDimensionWithVehicleCapacity(
                demand_callback_index, 0, [tractor_capacity] * num_vehicles, True, 'Capacity')

            # Search parameters
            search_parameters = pywrapcp.DefaultRoutingSearchParameters()
            search_parameters.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC
            search_parameters.local_search_metaheuristic = routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
            search_parameters.time_limit.seconds = 10
            for vehicle_id in range(num_vehicles):
                routing.SetFixedCostOfVehicle(10000, vehicle_id)

            solution = routing.SolveWithParameters(search_parameters)

            if solution:
                tractor_count = 1
                route_data = {
                    'Tractor': [],
                    'Stop Order': [],
                    'Name': [],
                    'Quantity (kg)': [],
                    'Utilization (%)': [],
                    'Google Maps Link': [],
                    'Highlight': []
                }

                for vehicle_id in range(num_vehicles):
                    index = routing.Start(vehicle_id)
                    route = []
                    load_this_route = 0
                    while not routing.IsEnd(index):
                        node_index = manager.IndexToNode(index)
                        route.append(node_index)
                        if node_index != 0:
                            load_this_route += demands[node_index]
                        index = solution.Value(routing.NextVar(index))
                    route.append(manager.IndexToNode(index))

                    if len(route) > 2:
                        utilization = (load_this_route / tractor_capacity) * 100
                        gmap_base = "https://www.google.com/maps/dir/"
                        waypoints = [f"{locations[i][0]},{locations[i][1]}" for i in route]
                        gmap_link = gmap_base + "/".join(waypoints)

                        for stop_index, i in enumerate(route):
                            name = 'Warehouse' if i == 0 else df.iloc[i - 1]['Supplier Name (Farmer Name)']
                            qty = 0 if i == 0 else df.iloc[i - 1]['Biomass Quantity (kg)']
                            util_str = f"{utilization:.1f}%" if stop_index == 0 else ""
                            highlight = utilization < 60 and stop_index == 0

                            route_data['Tractor'].append(f"Tractor {tractor_count}" if stop_index == 0 else "")
                            route_data['Stop Order'].append(stop_index + 1)
                            route_data['Name'].append(name)
                            route_data['Quantity (kg)'].append(qty)
                            route_data['Utilization (%)'].append(util_str)
                            route_data['Google Maps Link'].append(gmap_link if stop_index == 0 else "")
                            route_data['Highlight'].append(highlight)

                        tractor_count += 1

                # ✅ Show results in Streamlit (no Highlight column)
                route_df = pd.DataFrame(route_data)
                st.success(f"✅ Found {tractor_count - 1} optimized routes.")
                st.dataframe(route_df.drop(columns=['Highlight']))

                # 💾 Save to Excel with Yellow Highlight
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    route_df.drop(columns=['Highlight']).to_excel(writer, sheet_name=biomass_type, index=False)
                    workbook = writer.book
                    worksheet = writer.sheets[biomass_type]
                    yellow_format = workbook.add_format({'bg_color': '#FFFF00'})

                    for row_num, highlight in enumerate(route_df['Highlight'], start=1):  # Skip header
                        if highlight:
                            worksheet.set_row(row_num, None, yellow_format)

                st.download_button(
                    label="📥 Download Routes as Excel",
                    data=buffer.getvalue(),
                    file_name=f'{biomass_type}_routes.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            else:
                st.error("❌ No route solution found.")
    except Exception as e:
        st.error(f"⚠️ Error reading file: {e}")
else:
    st.info("ℹ️ Upload your supplier file to begin.")
