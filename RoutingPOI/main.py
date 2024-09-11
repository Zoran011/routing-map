import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import folium
from folium.plugins import AntPath
import openrouteservice
import webbrowser
import requests
from geopy.geocoders import Nominatim

class RoutingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Routing POI Application")

        self.create_widgets()
        self.geolocator_api_key = 'c12122fe22e54dbb8bf69844860f398d'  # Replace with your OpenCage API key
        self.client = openrouteservice.Client(key='5b3ce3597851110001cf624842bacdcbc44c4fef9c61772329f834be')  # Replace with your OpenRouteService API key
        self.poi_list = []
        self.df = None

    def create_widgets(self):
        self.load_button = tk.Button(self.root, text="Load Data", command=self.load_data)
        self.load_button.pack(pady=10)

        self.show_map_button = tk.Button(self.root, text="Show Map", command=self.show_map)
        self.show_map_button.pack(pady=10)

        self.save_map_button = tk.Button(self.root, text="Save Map with Route", command=self.save_map)
        self.save_map_button.pack(pady=10)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, orient='horizontal', length=300, mode='determinate', variable=self.progress_var)
        self.progress_bar.pack(pady=10)

    def geocode(self, address):
        try:
            response = requests.get(f'https://api.opencagedata.com/geocode/v1/json',
                                    params={'q': address, 'key': self.geolocator_api_key})
            result = response.json()
            if result['results']:
                location = result['results'][0]
                return (location['geometry']['lat'], location['geometry']['lng'])
            else:
                print(f"Geocoding failed for address: {address}")
                return None
        except Exception as e:
            print(f"Geocoding error: {e}")
            return None

    def load_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                self.poi_list = list(zip(self.df['Naziv'], self.df['Adresa']))
                messagebox.showinfo("Success", "Data loaded successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Error loading data: {e}")

    def show_map(self):
        if self.df is None or self.df.empty:
            messagebox.showerror("Error", "No data loaded.")
            return

        map_center = [44.8176, 20.4633]  # Center on Serbia
        m = folium.Map(location=map_center, zoom_start=7)

        geocoded_locations = []
        total = len(self.poi_list)
        for idx, (name, address) in enumerate(self.poi_list):
            self.progress_var.set((idx + 1) / total * 100)
            self.root.update_idletasks()
            coords = self.geocode(address)
            if coords:
                folium.Marker(location=coords, popup=f"{name}<br>{address}").add_to(m)
                geocoded_locations.append(coords)
            else:
                print(f"Failed to geocode address: {address}")

        self.progress_var.set(100)
        self.root.update_idletasks()

        if len(geocoded_locations) < 2:
            messagebox.showerror("Error", "Not enough locations to create a route.")
            return

        try:
            route = self.client.directions(
                coordinates=[list(reversed(coord)) for coord in geocoded_locations],
                profile='driving-car',
                format='geojson',
                optimize_waypoints=True
            )
            coords = route['features'][0]['geometry']['coordinates']
            if coords:
                coords = [list(reversed(coord)) for coord in coords]  # Convert to (lat, lon) format
                AntPath(locations=coords, dash_array=[20, 20]).add_to(m)
        except Exception as e:
            print(f"Error during routing: {e}")
            messagebox.showerror("Error", f"Error fetching route from OpenRouteService: {e}")

        # Add the user's current position using geolocation via JavaScript with HTTPS
        m.get_root().html.add_child(folium.Element("""
            <script>
            function addUserLocation(map) {
                if (navigator.geolocation) {
                    navigator.geolocation.getCurrentPosition(function(position) {
                        var lat = position.coords.latitude;
                        var lon = position.coords.longitude;
                        var userMarker = L.marker([lat, lon], {icon: L.icon({
                            iconUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-icon-red.png',
                            iconSize: [25, 41],
                            iconAnchor: [12, 41]
                        })}).addTo(map).bindPopup("You are here!");
                        map.setView([lat, lon], 13);
                    }, function() {
                        alert("Geolocation failed or was denied by the user.");
                    });
                } else {
                    alert("Geolocation is not supported by this browser.");
                }
            }
            addUserLocation(window.L);
            </script>
        """))

        map_file = 'map.html'
        m.save(map_file)
        webbrowser.open(map_file)

    def save_map(self):
        if self.df is None or self.df.empty:
            messagebox.showerror("Error", "No data loaded.")
            return

        map_center = [44.8176, 20.4633]  # Center on Serbia
        m = folium.Map(location=map_center, zoom_start=7)

        geocoded_locations = []
        for name, address in self.poi_list:
            coords = self.geocode(address)
            if coords:
                folium.Marker(location=coords, popup=f"{name}<br>{address}").add_to(m)
                geocoded_locations.append(coords)
            else:
                print(f"Failed to geocode address: {address}")

        if len(geocoded_locations) < 2:
            messagebox.showerror("Error", "Not enough locations to create a route.")
            return

        try:
            route = self.client.directions(
                coordinates=[list(reversed(coord)) for coord in geocoded_locations],
                profile='driving-car',
                format='geojson',
                optimize_waypoints=True
            )
            coords = route['features'][0]['geometry']['coordinates']
            if coords:
                coords = [list(reversed(coord)) for coord in coords]  # Convert to (lat, lon) format
                AntPath(locations=coords, dash_array=[20, 20]).add_to(m)
        except Exception as e:
            print(f"Error during routing: {e}")
            messagebox.showerror("Error", f"Error fetching route from OpenRouteService: {e}")

        # Add the user's current position using geolocation via JavaScript with HTTPS
        m.get_root().html.add_child(folium.Element("""
            <script>
            function addUserLocation(map) {
                if (navigator.geolocation) {
                    navigator.geolocation.getCurrentPosition(function(position) {
                        var lat = position.coords.latitude;
                        var lon = position.coords.longitude;
                        var userMarker = L.marker([lat, lon], {icon: L.icon({
                            iconUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-icon-red.png',
                            iconSize: [25, 41],
                            iconAnchor: [12, 41]
                        })}).addTo(map).bindPopup("You are here!");
                        map.setView([lat, lon], 13);
                    }, function() {
                        alert("Geolocation failed or was denied by the user.");
                    });
                } else {
                    alert("Geolocation is not supported by this browser.");
                }
            }
            addUserLocation(window.L);
            </script>
        """))

        save_file_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML files", "*.html")])
        if save_file_path:
            m.save(save_file_path)
            messagebox.showinfo("Success", f"Map with route saved as {save_file_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RoutingApp(root)
    root.mainloop()
