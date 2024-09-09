import requests

url = "http://devapigw.llogis.com:10100/api/address/print-info"
headers = {
    "Authorization": "IgtAK eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiJDMDEwODA2IiwiYXVkIjoiQzAxMDgwNiIsIm5hbWUiOiIzMDU4OTUiLCJleHAiOjE1MzUxMzU1OTk5OTksImlhdCI6MTcyNTg2NjE4M30.nBjlRs429bXrclVpJ-z3AX5sL0ekN82NQ94-P7WrlXk",
    "Content-Type": "application/json"  # Content-Type 헤더 추가
}

# 기본적으로 예상되는 파라미터
data = {
    "ID": "12345",           # 사용자 ID
    "network": "network1",    # 네트워크 이름
    "area_no": "123",         # 지역 번호
    "zip_no": "12345",        # 우편번호
    "address": "서울특별시"   # 주소
}

try:
    response = requests.post(url, headers=headers, json=data)
    response.raise_for_status()
    data = response.json()

    if "message" in data:
        print("결과 메시지:", data["message"])
    else:
        print("결과 메시지가 없습니다.")
except requests.exceptions.RequestException as e:
    print(f"API 요청 중 에러 발생: {e}")