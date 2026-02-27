import sys
import json
from typing import Callable
import packets.worker.main
import packets.calls.main


packets: dict[str, Callable[[dict[str, Callable]], None]] = {
    "workers": packets.worker.main.main,
    "calls": packets.calls.main.main,
}

def main():
    if len(sys.argv) < 2:
        raise TypeError

    filepath = sys.argv[1]

    with open(filepath, encoding='utf-8-sig') as json_file:
        excel_json = json.load(json_file)

    packet = excel_json.get("packet")
    if packet is None:
        print(excel_json)
        raise TypeError

    main_packet = packets.get(packet)
    if main_packet is None:
        print(packet)
        raise TypeError

    main_packet(excel_json)

if __name__ == "__main__":
    main()
