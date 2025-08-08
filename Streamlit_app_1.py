import streamlit as st
import pandas as pd
from geopy.distance import geodesic
from ortools.constraint_solver import pywrapcp, routing_enums_pb2
from io import BytesIO

st.set_page_config(page_title="Biomass Route Optimizer", layout="wide")
st.title("🚜 Biomass Route Optimizer")

# --- User Inputs ---
with st.expander("📍 Enter Warehouse Details"):
    warehouse_lat = st.number_input("Warehouse Latitude", format="%.6f", value=14.083970)
    warehouse_lon = st.number_input("Warehouse Longitude", format="%.6f", value=79.794420)

warehouse = ('Warehouse', warehouse_lat, warehouse_lon)

tractor_capacity = st.number_input("🛻 Tractor Capacity (kg):", min_value=100, step=100, value=2000)

# --- File Upload ---
st.subheader("📤 Upload Supplier Data (Google Sheet or Excel/CSV)")
uploaded_file = st.file_uploader("Upload File", type=['xlsx', 'csv'])

try:
    if uploaded_file.name.endswith('.csv'):
        df_full = pd.read_csv(uploaded_file)
    elif uploaded_file.name.endswith('.xlsx'):
        try:
            import openpyxl  # Try importing before using it
            df_full = pd.read_excel(uploaded_file, engine='openpyxl')
        except ImportError:
            st.error("❌ You need to install the `openpyxl` package to read Excel files.\n\nRun: `pip install openpyxl`")
            st.stop()
    else:
        st.error("❌ Unsupported file format. Please upload a .csv or .xlsx file.")
        st.stop()
except Exception as e:
    st.error(f"⚠️ Error reading file: {e}")
    st.stop()


        # Validate columns
        required_cols = {'Supplier Name (Farmer Name)', 'Latitude of the location', 'Longitude of the location', 'Biomass Type', 'Biomass Quantity'}
        if not required_cols.issubset(df_full.columns):
            st.error("❌ File must contain the columns: Supplier Name (Farmer Name), Latitude of the location, Longitude of the location, Biomass Type, Biomass Quantity")
        else:
            # Select biomass type
            biomass_type = st.selectbox("🌿 Select Biomass Type:", df_full['Type'].unique())

            if st.button("Generate Routes"):
                df = df_full[df_full['Type'] == biomass_type].reset_index(drop=True)
                locations = [(warehouse[1], warehouse[2])] + list(zip(df['Latitude'], df['Longitude']))
                demands = [0] + df['Quantity'].tolist()

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

                search_parameters = pywrapcp.DefaultRoutingSearchParameters()
                search_parameters.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC
                solution = routing.SolveWithParameters(search_parameters)

                if solution:
                    tractor_count = 1
                    route_data = {
                        'Tractor': [],
                        'Stop Order': [],
                        'Name': [],
                        'Quantity (kg)': [],
                        'Google Maps Link': []
                    }

                    for vehicle_id in range(num_vehicles):
                        index = routing.Start(vehicle_id)
                        route = []
                        route_load = 0
                        while not routing.IsEnd(index):
                            node_index = manager.IndexToNode(index)
                            route.append(node_index)
                            route_load += demands[node_index]
                            index = solution.Value(routing.NextVar(index))
                        route.append(manager.IndexToNode(index))

                        if len(route) > 2:
                            # Generate Google Maps link
                            gmap_base = "https://www.google.com/maps/dir/"
                            waypoints = [f"{locations[i][0]},{locations[i][1]}" for i in route]
                            gmap_link = gmap_base + "/".join(waypoints)

                            for stop_index, i in enumerate(route):
                                name = 'Warehouse' if i == 0 else df.iloc[i - 1]['Name']
                                qty = 0 if i == 0 else df.iloc[i - 1]['Quantity']

                                route_data['Tractor'].append(tractor_count if stop_index == 0 else "")
                                route_data['Stop Order'].append(stop_index + 1)
                                route_data['Name'].append(name)
                                route_data['Quantity (kg)'].append(qty)
                                route_data['Google Maps Link'].append(gmap_link if stop_index == 0 else "")

                            tractor_count += 1

                    route_df = pd.DataFrame(route_data)
                    st.success(f"✅ Found {tractor_count - 1} optimized routes.")
                    st.dataframe(route_df)

                    # Download Excel
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        route_df.to_excel(writer, sheet_name=biomass_type, index=False)
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
    st.info("ℹ️ Please upload a file to proceed.")
