import requests
import pandas as pd
from openpyxl.styles import PatternFill

with open('url_list.txt', 'r') as f:
    url_list = f.read().splitlines()


headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36"
}

status_description = {
    100: "Continue",
    101: "Switching Protocols",
    102: "Processing",
    103: "Early Hints",
    200: "OK",
    201: "Created",
    202: "Accepted",
    203: "Non-Authoritative Information",
    204: "No Content",
    205: "Reset Content",
    206: "Partial Content",
    207: "Multi-Status",
    208: "Already Reported",
    226: "IM Used",
    300: "Multiple Choices",
    301: "Moved Permanently",
    302: "Found",
    303: "See Other",
    304: "Not Modified",
    305: "Use Proxy",
    306: "Switch Proxy",
    307: "Temporary Redirect",
    308: "Permanent Redirect",
    400: "Bad Request",
    401: "Unauthorized",
    402: "Payment Required",
    403: "Forbidden",
    404: "Not Found",
    405: "Method Not Allowed",
    406: "Not Acceptable",
    407: "Proxy Authentication Required",
    408: "Request Timeout",
    409: "Conflict",
    410: "Gone",
    411: "Length Required",
    412: "Precondition Failed",
    413: "Payload Too Large",
    414: "URI Too Long",
    415: "Unsupported Media Type",
    416: "Range Not Satisfiable",
    417: "Expectation Failed",
    418: "I'm a teapot",
    421: "Misdirected Request",
    422: "Unprocessable Entity",
    423: "Locked",
    424: "Failed Dependency",
    425: "Too Early",
    426: "Upgrade Required",
    428: "Precondition Required",
    429: "Too Many Requests",
    431: "Request Header Fields Too Large",
    451: "Unavailable For Legal Reasons",
    500: "Internal Server Error",
    501: "Not Implemented",
    502: "Bad Gateway",
    503: "Service Unavailable",
    504: "Gateway Timeout",
    505: "HTTP Version Not Supported",
    506: "Variant Also Negotiates",
    507: "Insufficient Storage",
    508: "Loop Detected",
    510: "Not Extended",
    511: "Network Authentication Required",
}

result = []

for url in url_list:
    try:
        r = requests.get(url, headers=headers,allow_redirects=False)
        r.raise_for_status()
        if r.status_code in status_description:
            print("URL:", url)
            print("Status Code: {}".format(r.status_code))
            print("Status Description: {}".format(status_description[r.status_code]))
            # if status code start with 3 then add redirect url else add empty string
            if str(r.status_code).startswith('3'):
                print("Redirect URL: {}".format(r.headers['Location']))
                result.append([url, r.status_code, status_description[r.status_code], r.headers['Location']])
            else:
                print("Redirect URL: {}".format(''))
                result.append([url, r.status_code, status_description[r.status_code], ''])
        else:
            print("URL:", url)
            print("Status Code: {}".format(r.status_code))
            print("Status Description: {}".format("Unknown"))
            print("Redirect URL: {}".format(''))
            result.append([url, r.status_code, "Unknown", ''])
    except Exception as e:
        print(e)
        result.append([url, "Error", str(e), ''])

df = pd.DataFrame(result, columns=['URL', 'Status Code', 'Status Description', 'Redirect URL'])

writer = pd.ExcelWriter('result.xlsx', engine='openpyxl')

df.to_excel(writer, sheet_name='Sheet1', index=False)

wb = writer.book

ws = writer.sheets['Sheet1']

redFill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

for i in range(2, len(df) + 2):
    if df['Status Code'][i - 2] != 200:
        ws.cell(row=i, column=2).fill = redFill

writer.close()