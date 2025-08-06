import streamlit as st
import pandas as pd
from geopy.distance import geodesic
from ortools.constraint_solver import pywrapcp, routing_enums_pb2
from io import BytesIO

# --- Load Data ---
warehouse = ('Warehouse', 14.08397, 79.794420)
data = [
    ['John', 14.111326, 79.814047, 'Cotton', 900],
    ['Khan', 14.088685, 79.768994, 'Cotton', 800],
    ['Chan', 14.111699, 79.738138, 'Cotton', 1500],
    ['Saran', 14.093327, 79.75242,  'Cotton', 600],
    ['kiran', 14.10365,  79.787511, 'Cotton', 2000],
    ['F1',    14.128395, 79.763032, 'Cotton', 900],
    ['F2',    14.12099,  79.853734, 'Cotton', 1200],
    ['F3',    14.030065, 79.823203, 'Cotton', 1800],
    ['F4',    14.011394, 79.879007, 'Cotton', 1100],
    ['F5',    14.034023, 79.883283, 'Cotton', 1300],
    ['M1',    14.101598, 79.848863, 'Chilli', 750],
    ['M2',    14.11897,  79.857535, 'Chilli', 1450],
    ['M3',    14.171996, 79.79008,  'Chilli', 700],
    ['M4',    14.17054,  79.742417, 'Chilli', 800],
    ['M5',    14.119828, 79.733434, 'Chilli', 600],
    ['M6',    14.094221, 79.721397, 'Chilli', 1600],
]

df_full = pd.DataFrame(data, columns=['Name', 'Latitude', 'Longitude', 'Type', 'Quantity'])

# --- Streamlit UI ---
st.title("🚜 Biomass Route Optimizer")

biomass_type = st.selectbox("Select Biomass Type:", df_full['Type'].unique())
vehicle_capacity = st.number_input("Tractor Capacity (kg):", value=2000, step=100)

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
        demand_callback_index, 0, [vehicle_capacity] * num_vehicles, True, 'Capacity')

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

        # --- Download Excel ---
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
