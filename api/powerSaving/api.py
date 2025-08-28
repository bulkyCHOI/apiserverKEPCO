from ninja import NinjaAPI, Router

from django.http import JsonResponse, HttpResponse
import requests
import json
from datetime import date, timedelta, datetime
import pandas as pd

from io import BytesIO

api = NinjaAPI(csrf=False, docs_url='/docs/')

powerSaving_router = Router()
test_router = Router()
api.add_router("/test", test_router, tags=["Testing"])
api.add_router("/powerSaving", powerSaving_router, tags=["Power Saving"])

@test_router.get("/hello")
def hello(request, name:str="World"):
    return {"message": f"Hello {name}, Django Ninja!!!"}

@test_router.get("/datetime")
def get_datetime(request):
    now = datetime.now()
    return {"datetime": now.strftime("%Y-%m-%d %H:%M:%S")}

def convert_date_format(date_str):
    if len(date_str) == 8 and date_str.isdigit():
        return f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
    return "Invalid format"

def convert_time_format(time_str):
    if len(time_str) == 4 and time_str.isdigit():
        return f"{time_str[:2]}:{time_str[2:]}"
    return "Invalid format"
        
@powerSaving_router.get("/kepcoDailyData")
def kepcoDailyData(request, date:int=None, returnType:str="json"):
    """
    강북/강원 권역 고압국가 35개에 대해 일단위 합계 전력 사용량 데이터를 KEPCO API에서 가져옵니다.\n
    date 입력 포멧: YYYYMMDD(미입력시 어제 날짜).\n
    returnType 입력 포멧: json(기본값), xlsx.\n
    출력 형태는 아래와 같습니다.\n
    {
        "returnCode": "ok",
        "data": [
            {
                "Customer Number": "1234567890",
                "Date": "2024-10-01",
                "Bonbu": "Bonbu Name",
                "Center": "Center Name",
                "Team": "Team Name",
                "Guksa": "Guksa Name",
                "Power Usage": 12345.67
            },
            ...
        ]
    }
    If there is an error, the response will be:
    {
        "returnCode": "le",
        "error": "Error message"
    }
    """
    if (request.method == 'GET') :
        try:
            #Get the 'date' parameter from the request
            date = request.GET.get('date', None)

            # If 'date' is not provided, use yesterday's date
            if date is None:
                today = datetime.today() - timedelta(days=1)
                date = today.strftime("%Y%m%d")  # Format as YYYYMMDD
            else:
                # Ensure the date is in string format
                date = str(date)

            # Input CSV file path
            input_csv = "kepcolist_gg.csv"

            # API URL and static parameters
            # https://opm.kepco.co.kr:11080/OpenAPI/getDayLpData.do?custNo=0135338560&date=20241001&serviceKey=bpb89eyd7bg430vckh8t&returnType=02
            url = "https://opm.kepco.co.kr:11080/OpenAPI/getDayLpData.do"
            service_key = "bpb89eyd7bg430vckh8t"

            # Read customer numbers from CSV
            customer_data = pd.read_csv(input_csv, dtype=str)

            # Prepare a list to store the results
            results = []

            for _, row in customer_data.iterrows():
                cust_no = row.get("고객번호")
                if not cust_no:
                    continue
                bonbu = row.get("본부명")
                center = row.get("센터")
                team = row.get("팀")
                guksa = row.get("국사")

                # API call
                params = {
                    "custNo": cust_no,
                    "date": date,
                    "serviceKey": service_key,
                    "returnType": "02"
                }
                response = requests.get(url, params=params)

                if response.status_code == 200:
                    data = response.json()
                    day_lp_data = data.get("dayLpDataInfoList", [])

                    if day_lp_data:
                        total_power = 0
                        for record in day_lp_data:
                            total_power += sum(
                                value for key, value in record.items() if key.startswith("pwr_qty") and isinstance(value, (int, float))
                            )
                        results.append({
                            "Customer Number": cust_no,
                            "Date": convert_date_format(date),
                            "Bonbu": bonbu,
                            "Center": center,
                            "Team": team,
                            "Guksa": guksa,
                            "Power Usage": total_power
                        })
                    else:
                        results.append({
                            "Customer Number": cust_no,
                            "Date": convert_date_format(date),
                            "Bonbu": bonbu,
                            "Center": center,
                            "Team": team,
                            "Guksa": guksa,
                            "Power Usage": None
                        })
                else:
                    results.append({
                        "Customer Number": cust_no,
                        "Date": convert_date_format(date),
                        "Bonbu": bonbu,
                        "Center": center,
                        "Team": team,
                        "Guksa": guksa,
                        "Power Usage": "API Error"
                    })

            # Convert results to a DataFrame
            df = pd.DataFrame(results)

            if returnType == "json":
                # Convert DataFrame to JSON
                json_result = df.to_json(orient="records", force_ascii=False)

                # Return JSON response
                return JsonResponse({"returnCode": "ok", "data": json.loads(json_result)}, json_dumps_params={'ensure_ascii': False})
            elif returnType == "xlsx":
                # Create a BytesIO buffer to store the Excel file
                output = BytesIO()
                # Convert DataFrame to Excel
                df.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)

                # Return Excel file as a downloadable response
                response = HttpResponse(
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    content=output.getvalue()
                )
                response['Content-Disposition'] = f'attachment; filename="kepco_daily_data_{date}.xlsx"'
                return response
            else:
                return JsonResponse({"returnCode": "le", "error": "Invalid returnType. Use 'json' or 'xlsx'."}, json_dumps_params={'ensure_ascii': False})

        except Exception as e:
            print(str(e))
            return JsonResponse({"returnCode": "le", "error": str(e)}, json_dumps_params={'ensure_ascii': False})
        
@powerSaving_router.get("/kepcoDailyData15min")
def kepcoDailyData15min(request, date:int=None, returnType:str="json"):
    """
    강북/강원 권역 고압국가 35개에 대해 일단위 15분간 전력 사용량 데이터를 KEPCO API에서 가져옵니다.\n
    date 입력 포멧: YYYYMMDD(미입력시 어제 날짜).\n
    returnType 입력 포멧: json(기본값), xlsx.\n
    출력 형태는 아래와 같습니다.\n
    {
        "returnCode": "ok",
        "data": [
            {
                "Customer Number": "1234567890",
                "MeterNo": "Meter123",
                "Date": "2024-10-01",
                "Time": "00:15",
                "Bonbu": "Bonbu Name",
                "Center": "Center Name",
                "Team": "Team Name",
                "Guksa": "Guksa Name",
                "Power Usage": 123.45
            },
            ...
        ]
    }
    If there is an error, the response will be:
    {
        "returnCode": "le",
        "error": "Error message"
    }
    """
    if (request.method == 'GET') :
        try:
            #Get the 'date' parameter from the request
            date = request.GET.get('date', None)

            # If 'date' is not provided, use yesterday's date
            if date is None:
                today = datetime.today() - timedelta(days=1)
                date = today.strftime("%Y%m%d")  # Format as YYYYMMDD
            else:
                # Ensure the date is in string format
                date = str(date)
                
            # Input CSV file path
            input_csv = "kepcolist_gg.csv"

            # API URL and static parameters
            url = "https://opm.kepco.co.kr:11080/OpenAPI/getDayLpData.do"
            service_key = "bpb89eyd7bg430vckh8t"
    
            # Read customer numbers from CSV
            customer_data = pd.read_csv(input_csv, dtype=str)

            # Prepare a list to store the results
            results = []

            for _, row in customer_data.iterrows():
                cust_no = row.get("고객번호")
                if not cust_no:
                    continue
                bonbu = row.get("본부명")
                center = row.get("센터")
                team = row.get("팀")
                guksa = row.get("국사")

                # API call
                params = {
                    "custNo": cust_no,
                    "date": date,
                    "serviceKey": service_key,
                    "returnType": "02"
                }
                response = requests.get(url, params=params)

                if response.status_code == 200:
                    data = response.json()
                    day_lp_data = data.get("dayLpDataInfoList", [])

                    if day_lp_data:
                        for record in day_lp_data:
                            # pwr_qtyXXXX 형식의 키를 순회
                            for key, value in record.items():
                                if key.startswith("pwr_qty") and isinstance(value, (int, float)):
                                    # 키에서 시간 추출 (예: pwr_qty0015 → 00:15)
                                    time_str = key[-4:]  # 마지막 4자리 (HHMM)
                                    results.append({
                                        "Customer Number": cust_no,
                                        "MeterNo": record.get("meterNo"),
                                        "Date": convert_date_format(date),
                                        "Time": time_str,
                                        "Bonbu": bonbu,
                                        "Center": center,
                                        "Team": team,
                                        "Guksa": guksa,
                                        "Power Usage": value
                                    })
                    else:
                        return JsonResponse({
                            "returnCode": "le",
                            "error": f"No data found for Customer Number {cust_no}"
                        }, json_dumps_params={'ensure_ascii': False})
                else:
                    return JsonResponse({
                        "returnCode": "le",
                        "error": f"API Error for Customer Number {cust_no}: {response.status_code}"
                    }, json_dumps_params={'ensure_ascii': False})

            # Convert results to a DataFrame
            df = pd.DataFrame(results)

            if returnType == "json":
                # Convert DataFrame to JSON
                json_result = df.to_json(orient="records", force_ascii=False)

                # Return JSON response
                return JsonResponse({"returnCode": "ok", "data": json.loads(json_result)}, json_dumps_params={'ensure_ascii': False})
            elif returnType == "xlsx":
                # Create a BytesIO buffer to store the Excel file
                output = BytesIO()
                # Convert DataFrame to Excel
                df.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)

                # Return Excel file as a downloadable response
                response = HttpResponse(
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    content=output.getvalue()
                )
                response['Content-Disposition'] = f'attachment; filename="kepco_daily_15min_data_{date}.xlsx"'
                return response
            else:
                return JsonResponse({"returnCode": "le", "error": "Invalid returnType. Use 'json' or 'xlsx'."}, json_dumps_params={'ensure_ascii': False})

        except Exception as e:
            print(str(e))
            return JsonResponse({"returnCode": "le", "error": str(e)}, json_dumps_params={'ensure_ascii': False})
        
@powerSaving_router.get("/kepco15minData")
def kepco15minData(request, dateTime:int, returnType:str="json"):
    """
    강북/강원 권역 고압국가 35개에 대해 15분간 전력 사용량 데이터를 KEPCO API에서 가져옵니다.\n
    dateTime 입력 포멧: YYYYMMDDHHMM.\n
    returnType 입력 포멧: json(기본값), xlsx.\n
    출력 형태는 아래와 같습니다.\n
    {
        "returnCode": "ok",
        "data": [
            {
                "Customer Number": "1234567890",
                "MeterNo": "Meter123",
                "Date": "2024-10-01",
                "Time": "00:15",
                "Bonbu": "Bonbu Name",
                "Center": "Center Name",
                "Team": "Team Name",
                "Guksa": "Guksa Name",
                "Power Usage": 123.45
            },
            ...
        ]
    }
    If there is an error, the response will be:
    {
        "returnCode": "le",
        "error": "Error message"
    }   
    """
    if (request.method == 'GET') :
        try:
            # Input CSV file path
            input_csv = "kepcolist_gg.csv"

            # API URL and static parameters
            url = "https://opm.kepco.co.kr:11080/OpenAPI/getMinuteLpData.do"
            service_key = "bpb89eyd7bg430vckh8t"

            # Read customer numbers from CSV
            customer_data = pd.read_csv(input_csv, dtype=str)

            # Prepare a list to store the results
            results = []

            for _, row in customer_data.iterrows():
                cust_no = row.get("고객번호")
                if not cust_no:
                    continue
                bonbu = row.get("본부명")
                center = row.get("센터")
                team = row.get("팀")
                guksa = row.get("국사")

                # API call
                params = {
                    "custNo": cust_no,
                    "dateTime": dateTime,
                    "serviceKey": service_key,
                    "returnType": "02"
                }
                response = requests.get(url, params=params)

                if response.status_code == 200:
                    data = response.json()
                    minute_lp_data = data.get("minuteLpDataInfoList", [])

                    if minute_lp_data:
                        for record in minute_lp_data:
                            for key, value in record.items():
                                if key=="pwr_qty" and isinstance(value, (int, float)):
                                    results.append({
                                        "Customer Number": cust_no,
                                        "MeterNo": record.get("meterNo"),
                                        "Date": convert_date_format(record.get("mr_ymd")),
                                        "Time": record.get("mr_hhmi"),
                                        "Bonbu": bonbu,
                                        "Center": center,
                                        "Team": team,
                                        "Guksa": guksa,
                                        "Power Usage": value
                                    })
                    else:
                        return JsonResponse({
                            "returnCode": "le",
                            "error": f"No data found for Customer Number {cust_no}"
                        }, json_dumps_params={'ensure_ascii': False})
                else:
                    return JsonResponse({
                        "returnCode": "le",
                        "error": f"API Error for Customer Number {cust_no}: {response.status_code}"
                    }, json_dumps_params={'ensure_ascii': False})

            # Convert results to a DataFrame
            df = pd.DataFrame(results)

            if returnType == "json":
                # Convert DataFrame to JSON
                json_result = df.to_json(orient="records", force_ascii=False)

                # Return JSON response
                return JsonResponse({"returnCode": "ok", "data": json.loads(json_result)}, json_dumps_params={'ensure_ascii': False})
            elif returnType == "xlsx":
                # Create a BytesIO buffer to store the Excel file
                output = BytesIO()
                # Convert DataFrame to Excel
                df.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)

                # Return Excel file as a downloadable response
                response = HttpResponse(
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    content=output.getvalue()
                )
                response['Content-Disposition'] = f'attachment; filename="kepco_15min_data_{dateTime}.xlsx"'
                return response
            else:
                return JsonResponse({"returnCode": "le", "error": "Invalid returnType. Use 'json' or 'xlsx'."}, json_dumps_params={'ensure_ascii': False})
        except Exception as e:
            print(str(e))
            return JsonResponse({"returnCode": "le", "error": str(e)}, json_dumps_params={'ensure_ascii': False})

@powerSaving_router.get("/kepcoDailyRangeData")
def kepcoDailyRangeData(request, startDate: int = None, endDate: int = None, returnType: str = "json"):
    """
    강북/강원 권역 고압국가 35개에 대해 기간 내 일단위 합계 전력 사용량 데이터를 KEPCO API에서 가져옵니다.
    startDate, endDate 입력 포멧: YYYYMMDD (미입력시 어제~어제).
    returnType 입력 포멧: json(기본값), xlsx.
    """
    if request.method == 'GET':
        try:
            # 날짜 파라미터 처리
            if startDate is None or endDate is None:
                yesterday = datetime.today() - timedelta(days=1)
                startDate = endDate = yesterday.strftime("%Y%m%d")
            else:
                startDate = str(startDate)
                endDate = str(endDate)

            # 날짜 리스트 생성
            start_dt = datetime.strptime(startDate, "%Y%m%d")
            end_dt = datetime.strptime(endDate, "%Y%m%d")
            date_list = [(start_dt + timedelta(days=i)).strftime("%Y%m%d") for i in range((end_dt - start_dt).days + 1)]

            input_csv = "kepcolist_gg.csv"
            url = "https://opm.kepco.co.kr:11080/OpenAPI/getDayLpData.do"
            service_key = "bpb89eyd7bg430vckh8t"
            customer_data = pd.read_csv(input_csv, dtype=str)
            results = []

            for date in date_list:
                for _, row in customer_data.iterrows():
                    cust_no = row.get("고객번호")
                    if not cust_no:
                        continue
                    bonbu = row.get("본부명")
                    center = row.get("센터")
                    team = row.get("팀")
                    guksa = row.get("국사")

                    params = {
                        "custNo": cust_no,
                        "date": date,
                        "serviceKey": service_key,
                        "returnType": "02"
                    }
                    response = requests.get(url, params=params)

                    if response.status_code == 200:
                        data = response.json()
                        day_lp_data = data.get("dayLpDataInfoList", [])
                        if day_lp_data:
                            total_power = 0
                            for record in day_lp_data:
                                total_power += sum(
                                    value for key, value in record.items() if key.startswith("pwr_qty") and isinstance(value, (int, float))
                                )
                            results.append({
                                "Customer Number": cust_no,
                                "Date": convert_date_format(date),
                                "Bonbu": bonbu,
                                "Center": center,
                                "Team": team,
                                "Guksa": guksa,
                                "Power Usage": total_power
                            })
                        else:
                            results.append({
                                "Customer Number": cust_no,
                                "Date": convert_date_format(date),
                                "Bonbu": bonbu,
                                "Center": center,
                                "Team": team,
                                "Guksa": guksa,
                                "Power Usage": None
                            })
                    else:
                        results.append({
                            "Customer Number": cust_no,
                            "Date": convert_date_format(date),
                            "Bonbu": bonbu,
                            "Center": center,
                            "Team": team,
                            "Guksa": guksa,
                            "Power Usage": "API Error"
                        })

            df = pd.DataFrame(results)

            if returnType == "json":
                json_result = df.to_json(orient="records", force_ascii=False)
                return JsonResponse({"returnCode": "ok", "data": json.loads(json_result)}, json_dumps_params={'ensure_ascii': False})
            elif returnType == "xlsx":
                output = BytesIO()
                df.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)
                response = HttpResponse(
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    content=output.getvalue()
                )
                response['Content-Disposition'] = f'attachment; filename="kepco_daily_range_data_{startDate}_{endDate}.xlsx"'
                return response
            else:
                return JsonResponse({"returnCode": "le", "error": "Invalid returnType. Use 'json' or 'xlsx'."}, json_dumps_params={'ensure_ascii': False})

        except Exception as e:
            print(str(e))
            return JsonResponse({"returnCode": "le", "error": str(e)}, json_dumps_params={'ensure_ascii': False})