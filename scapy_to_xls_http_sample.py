#!/usr/bin/env python
# coding: shift-jis

import re
from scapy.all import *
from datetime import datetime
from excel_wrapper import ExcelWapper


def get_http_request_or_response(keyword, contents):
    """ Get HTTP request or response text data

    ex. 
    contents -> "b'HTTP/1.1 200 OK\r\n.........."
    keyword  -> "HTTP"
    output   -> "HTTP/1.1 200 OK\r\n"

    Parameters
    ----------

    keyword : str
        POST, GET, HTTP ... etc

    contents: str
        str(packet['Raw'])
        packet is rdpcap(name).filter[n]

    Returns
    -------

    text : str
        HTTP request or response text data

    """
    pattern = r".*?(" + keyword + r").*?(\\r\\n)"
    match = re.match(pattern, contents)
    if match == None:
        return ""

    # ex. "'bGET /" -> "GET /"
    offset = match.group(0).find(keyword)
    text = match.group(0)[offset:]
    return text


def find_http_data(packet):
    """ Find HTTP data from packet data.

    If the packet include  http request or response , then pick up its text.
    (Request GET/POST, Response HTTP)

    ex 
    "b'GET / \r\n..................." -> "GET / \r\n"
    "b'HTTP/1.1 200 OK\r\n.........." -> "HTTP/1.1 200 OK\r\n"
    "b'HOGE / \r\n.................." -> None

    Parameters
    ----------

    packet : scapy.layers.l2.Ether
        packet is rdpcap(name).filter(filter)[n]

    Returns
    -------

    text : str
        HTTP request or response text data

    """
    payload = str(packet['Raw'])
    request = get_http_request_or_response(r"POST", payload)
    request += get_http_request_or_response(r"GET", payload)
    if len(request) > 0:
        return request
    else:
        response = get_http_request_or_response(r"HTTP", payload)
        if len(response) > 0:
            return response
        else:
            return None


def analyze_http_captured_file(name, port):
    """ Analyze capture file(*.pcap) as HTTP transfferd.

    This function outputs  list data. like this,

    [(timestump, summary, http_text), (timestump, summary, http_text), ...]

    touple example.
        ( 2019-03-28T07:49:11.13361,   # index 0
          Ether / IP / TCP 192.168.0.11:49875 > 192.168.0.12:http PA / Raw, #index 1
          GET / HTTP/1.0\r\n ) # index 2

    Parameters
    ----------

    name : str 
        pcap file name

    port : int
        HTTP port number

    Returns
    -------

    text : list
        report data
        [(timestump, summary, http_text), (timestump, summary, http_text), ...]

    """
    def filter(p): return Raw in p and TCP in p and (
        p[TCP].sport == port or p[TCP].dport == port)
    packets = rdpcap(name).filter(filter)
    http_list = []

    for packet in packets:
        text = find_http_data(packet)
        if text != None:
            datetime_text = datetime.fromtimestamp(packet.time).isoformat()
            http_list.append((datetime_text, packet.summary(), text))

    return http_list


def print_http_list(http_list):
    """ Print analyze_http_captured_file function's result.

    Parameters
    ----------

    http_list : list
        analyze_http_captured_file function's result

    """
    i = 1
    for http_item in http_list:
        datetime_text = http_item[0]
        summary_text = http_item[1]
        http_text = http_item[2]
        print("No:", i, " ", datetime_text)
        print("\t", summary_text)
        print("\t", http_text)
        i += 1


def write_http_data_to_excel(excel, x, y, http_item):
    """ Write analyze_http_captured_file function's result[n] to excel 

    Parameters
    ----------

    excel : objet
        ExcelWapper object

    x : int
        sheet position x

    y : int
        sheet position y

    http_item : taple
        taple of analyze_http_captured_file function's result
        ex. 
            result = analyze_http_captured_file(name, file)
            write_http_data_to_excel(excel, x, y, result[n]) # <- like this

    """
    datetime_text = http_item[0]
    summary_text = http_item[1]
    http_text = http_item[2]
    excel.write_value(x + 0,  y, datetime_text)
    excel.write_value(x + 1,  y, summary_text)
    excel.write_value(x + 2,  y, http_text)


def make_excel_file(http_list, filename, img_filename):
    """ Make report excel file

    This function makes a excel file with using ExcelWapper class.
    If you want to know detail, see excel_wrapper.py(docstirng)

    Parameters
    ----------

    http_list : list
        analyze_http_captured_file function's result

    filename : str
        Filename of excel.

    filename : str
        Filename of image(jpg / png etc...).

    """

    sheet_name = "ScapyResut"
    excel = ExcelWapper()
    excel.create_book()
    excel.create_sheet(sheet_name)

    x_start = x_pos = 2
    x_size = 3
    y_start = y_pos = 4

    # title
    http_item = ("TimeStump", "Summary", "HTTP data")
    write_http_data_to_excel(excel, x_pos, y_pos, http_item)
    y_pos += 1

    # values
    for http_item in http_list:
        write_http_data_to_excel(excel, x_pos, y_pos, http_item)
        y_pos += 1

    excel.resize_sheet_width()
    excel.draw_table(x_start, x_size, y_start, (y_pos - y_start))

    # image
    if img_filename != None:
        excel.add_imagefile(2, 2, img_filename, False, True)

    # save file
    excel.save(filename)

    print("scapy_to_xls_httpsample:", filename, " created.")


if __name__ == "__main__":
    """ Sample program of scapy and openpyxl
    """

    # arguments
    if len(sys.argv) < 3:
        print(
            "python3x scapy_to_xls_httpsampple.py [pcap filename] [port] [excel filename(option)]")
        sys.exit(1)
    img_filename = "report_sample.png"
    excel_filename = None
    pcap_filename = sys.argv[1]
    http_port = int(sys.argv[2])
    if len(sys.argv) == 4:
        excel_filename = sys.argv[3]
    print("scapy_to_xls_httpsample: pcap:", pcap_filename,
          " port:", http_port, "excel:", excel_filename)

    # scapy
    http_list = analyze_http_captured_file(pcap_filename, http_port)
    print_http_list(http_list)

    # openpyxl
    if excel_filename != None:
        make_excel_file(http_list, excel_filename, img_filename)

    print("scapy_to_xls_httpsample:completed.")
