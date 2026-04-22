import requests
import polyline
from geopy.distance import geodesic


class RouteEngine:
    @staticmethod
    def get_full_road_route(coords_tuple: tuple):
        """Calcula rota via OSRM com fallback para linha reta."""
        coords = list(coords_tuple)
        if len(coords) < 2:
            return [], 0, 0

        full_geometry = []
        total_duration = 0
        total_distance = 0
        chunk_size = 90

        for i in range(0, len(coords) - 1, chunk_size - 1):
            chunk = coords[i:i + chunk_size]
            if len(chunk) < 2:
                break

            path_segments = ";".join(f"{c[1]},{c[0]}" for c in chunk)
            url = (
                "http://router.project-osrm.org/route/v1/driving/"
                f"{path_segments}?overview=full&geometries=polyline"
            )

            try:
                response = requests.get(url, timeout=10).json()
                if response.get("code") == "Ok":
                    route = response["routes"][0]
                    decoded = polyline.decode(route["geometry"])
                    if full_geometry and decoded:
                        full_geometry.extend(decoded[1:])
                    else:
                        full_geometry.extend(decoded)
                    total_duration += route["duration"]
                    total_distance += route["distance"]
                    continue
            except Exception:
                pass

            # Fallback: aproxima por linha reta a ~40 km/h.
            for j in range(len(chunk) - 1):
                dist_reta = geodesic(chunk[j], chunk[j + 1]).meters
                total_distance += dist_reta
                total_duration += dist_reta / 11.1
                if full_geometry:
                    full_geometry.append(chunk[j + 1])
                else:
                    full_geometry.extend([chunk[j], chunk[j + 1]])

        return full_geometry, total_duration, total_distance

    @staticmethod
    def auto_optimize(points):
        """Nearest neighbor a partir da sede (indice 0)."""
        n = len(points)
        if n < 2:
            return list(range(n))

        unvisited = list(range(1, n))
        current = 0
        route = [0]

        while unvisited:
            nearest = min(
                unvisited,
                key=lambda i: geodesic(points[current]["coord"], points[i]["coord"]).meters,
            )
            route.append(nearest)
            unvisited.remove(nearest)
            current = nearest

        return route


def gerar_cores_gradiente(n):
    """Gera n cores em gradiente (azul -> roxo -> vermelho)."""
    if n <= 1:
        return ["#0066CC"]

    cores = []
    for i in range(n):
        ratio = i / (n - 1)
        if ratio < 0.5:
            r = int(0 * (1 - ratio * 2) + 128 * ratio * 2)
            g = int(102 * (1 - ratio * 2) + 0 * ratio * 2)
            b = int(204 * (1 - ratio * 2) + 204 * ratio * 2)
        else:
            r_ratio = (ratio - 0.5) * 2
            r = int(128 * (1 - r_ratio) + 255 * r_ratio)
            g = 0
            b = int(204 * (1 - r_ratio))
        cores.append(f"#{r:02x}{g:02x}{b:02x}")
    return cores


class DataMaster:
    @staticmethod
    def fix_coord(val):
        """Normaliza latitude/longitude."""
        try:
            if val is None:
                return None

            if isinstance(val, str):
                val = val.strip()
                if not val:
                    return None
                val = val.replace(" ", "")
                if "," in val and "." in val:
                    val = val.replace(".", "").replace(",", ".")
                elif "," in val:
                    val = val.replace(",", ".")

            val_float = float(val)
            if val_float == 0:
                return None
            if abs(val_float) > 1000:
                val_float = val_float / 1_000_000.0
            if -180 <= val_float <= 180:
                return round(val_float, 8)
            return None
        except Exception:
            return None
