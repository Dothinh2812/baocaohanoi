import pandas as pd
import math
import json

def haversine_distance(lat1, lon1, lat2, lon2):
    """
    Tính khoảng cách giữa hai điểm trên Trái Đất sử dụng công thức Haversine
    
    Args:
        lat1, lon1: Vĩ độ và kinh độ của điểm 1 (độ)
        lat2, lon2: Vĩ độ và kinh độ của điểm 2 (độ)
    
    Returns:
        Khoảng cách tính bằng km
    """
    # Chuyển đổi từ độ sang radian
    lat1_rad = math.radians(lat1)
    lon1_rad = math.radians(lon1)
    lat2_rad = math.radians(lat2)
    lon2_rad = math.radians(lon2)
    
    # Tính hiệu số
    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad
    
    # Công thức Haversine
    a = math.sin(dlat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon/2)**2
    c = 2 * math.asin(math.sqrt(a))
    
    # Bán kính Trái Đất (km)
    earth_radius = 6371
    
    # Tính khoảng cách
    distance = earth_radius * c
    
    return distance

def find_nearest_bts_station(user_lat, user_long, cabinet_name, excel_file_path="ket_qua_gop.xlsx"):
    """
    Tìm khoảng cách từ hộp (cabinet_name) đến tọa độ nhận được
    Args:
        user_lat: Vĩ độ nhận được từ OpenAI
        user_long: Kinh độ nhận được từ OpenAI
        cabinet_name: Tên hộp (ví dụ H-BVI/2024)
        excel_file_path: Đường dẫn đến file Excel chứa dữ liệu hộp
    Returns:
        Dictionary chứa thông tin hộp và khoảng cách
    """
    try:
        # Đọc dữ liệu từ file Excel
        df = pd.read_excel(excel_file_path)

        # Kiểm tra các cột cần thiết
        required_columns = ['Tên kết cuối', 'Vĩ độ', 'Kinh độ']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Không tìm thấy cột '{col}' trong file Excel")

        # Chuyển đổi tọa độ nhận được sang float
        user_lat = float(user_lat)
        user_long = float(user_long)

        print(f"Tìm hộp {cabinet_name} trong file {excel_file_path}")

        # Tìm dòng chứa tên hộp
        matched_row = None
        for idx, row in df.iterrows():
            if cabinet_name in str(row['Tên kết cuối']):
                matched_row = row
                break

        if matched_row is None:
            return {
                'error': f'Không tìm thấy hộp {cabinet_name} trong file',
                'user_coordinates': {'lat': user_lat, 'long': user_long}
            }

        # Lấy tọa độ của hộp từ file
        try:
            box_lat = float(matched_row['Vĩ độ'])
            box_long = float(matched_row['Kinh độ'])
        except Exception as e:
            return {
                'error': f'Lỗi lấy tọa độ hộp: {str(e)}',
                'user_coordinates': {'lat': user_lat, 'long': user_long}
            }

        # Tính khoảng cách từ tọa độ nhận được đến hộp
        distance_km = haversine_distance(user_lat, user_long, box_lat, box_long)

        result = {
            'user_coordinates': {
                'lat': user_lat,
                'long': user_long
            },
            'cabinet': {
                'cabinet_name': cabinet_name,
                'box_lat': box_lat,
                'box_long': box_long
            },
            'distance_km': round(distance_km, 3)
        }

        print(f"Khoảng cách từ tọa độ nhận được đến hộp {cabinet_name}: {distance_km:.3f} km")
        return result

    except FileNotFoundError:
        return {
            'error': f'Không tìm thấy file Excel: {excel_file_path}',
            'user_coordinates': {'lat': user_lat, 'long': user_long}
        }
    except Exception as e:
        return {
            'error': f'Lỗi xử lý: {str(e)}',
            'user_coordinates': {'lat': user_lat, 'long': user_long}
        }

def get_nearest_bts_name_only(user_lat, user_long, excel_file_path="map_gps_all_bts.xlsx"):
    """
    Trả về chỉ tên trạm BTS gần nhất (để tích hợp vào webhook)
    
    Args:
        user_lat: Vĩ độ của người dùng
        user_long: Kinh độ của người dùng
        excel_file_path: Đường dẫn đến file Excel chứa dữ liệu BTS
    
    Returns:
        String: Tên trạm BTS gần nhất hoặc thông báo lỗi
    """
    result = find_nearest_bts_station(user_lat, user_long, excel_file_path)
    
    if 'error' in result:
        return f"Lỗi: {result['error']}"
    
    return result['nearest_station']['TEN_TRAM']

def get_nearest_bts_with_distance(user_lat, user_long, excel_file_path="map_gps_all_bts.xlsx"):
    """
    Trả về tên trạm BTS gần nhất và khoảng cách (để tích hợp vào webhook)
    
    Args:
        user_lat: Vĩ độ của người dùng
        user_long: Kinh độ của người dùng
        excel_file_path: Đường dẫn đến file Excel chứa dữ liệu BTS
    
    Returns:
        Tuple: (Tên trạm BTS gần nhất, Khoảng cách (m)) hoặc (thông báo lỗi, None)
    """
    result = find_nearest_bts_station(user_lat, user_long, excel_file_path)
    
    if 'error' in result:
        return (f"Lỗi: {result['error']}", None)
    
    # Chuyển đổi từ km sang m (1 km = 1000 m)
    distance_km = result['nearest_station']['distance_km']
    distance_m = round(distance_km * 1000)
    
    return (
        result['nearest_station']['TEN_TRAM'], 
        distance_m
    )

# Test function
if __name__ == "__main__":
    # Test với tọa độ mẫu
    test_lat = 21.09556205
    test_long = 105.32388554
    
    print("=" * 60)
    print("TEST TÍNH TOÁN KHOẢNG CÁCH BTS")
    print("=" * 60)
    
    # Test function đầy đủ với cabinet_name
    test_cabinet_name = "H-BVI/2024"  # Thay bằng tên hộp thực tế để test
    result = find_nearest_bts_station(test_lat, test_long, test_cabinet_name)
    print("\nKết quả đầy đủ:")
    print(json.dumps(result, indent=2, ensure_ascii=False))

    print("\n" + "=" * 60)