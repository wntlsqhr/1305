import requests
import json

# API 요청에 사용할 URL과 API 키
url = "http://devapigw.llogis.com:10100/api/address/print-info"
api_key = "eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiJDMDEwODA2IiwiYXVkIjoiQzAxMDgwNiIsIm5hbWUiOiIzMDU4OTUiLCJleHAiOjE1MzUxMzU1OTk5OTksImlhdCI6MTcyNTg2NjE4M30.nBjlRs429bXrclVpJ-z3AX5sL0ekN82NQ94-P7WrlXk"  # 여기에 실제 API 키를 입력하세요.

# 요청에 필요한 헤더 설정
headers = {
    "Authorization": f"IgtAK eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiJDMDEwODA2IiwiYXVkIjoiQzAxMDgwNiIsIm5hbWUiOiIzMDU4OTUiLCJleHAiOjE1MzUxMzU1OTk5OTksImlhdCI6MTcyNTg2NjE4M30.nBjlRs429bXrclVpJ-z3AX5sL0ekN82NQ94-P7WrlXk",
    "Content-Type": "application/json"  # JSON 데이터를 보내기 위한 Content-Type 헤더
}

# 요청에 필요한 필터 값 설정 (JSON 형식)
filter_value = {
    "zip_no": "153802",
    "address": "서울 금천구 가산동 345-1테스트(디지털로 154)",
    "area_no": "",
    "network": "00",
    "ID": ""
}

# GET 요청 보내기 (필터 값을 JSON으로 전달)
response = requests.get(url, headers=headers, params={"filter": json.dumps(filter_value)}, timeout=10)

# 응답 데이터 확인
if response.status_code == 200:
    # JSON 형식의 응답 데이터를 파싱하여 출력
    data = response.json()
    print("Response Data:")
    for key, value in data.items():
        print(f"{key}: {value}")
else:
    # 오류 발생 시 상태 코드와 메시지 출력
    print(f"Error {response.status_code}: {response.text}")