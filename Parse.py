
def decode_key(text):
    return ''.join([i if ord(i) < 128 else '<?>' for i in text])


def big_endian(data):
    big = ""
    for i in range(0, len(data), 4):
        big += data[i+2:i+4]
        big += data[i:i+2]
    return big


def parse_date(d):
    d = big_endian(d)
    year = int(d[0:4], 16)
    month = int(d[4:8], 16)
    day = int(d[12:16], 16)

    hour = int(d[16:20], 16)
    minutes = int(d[20:24], 16)
    seconds = int(d[24:28], 16)
    return str(year) + "-" + pad(month) + "-" + pad(day) + " " + pad(hour) + ":" + pad(minutes) + ":" + pad(seconds)
    # return None


def pad(data):
    if len(str(data)) != 2:
        return "0" + str(data)
    return str(data)


def hex_windows_to_date(dt):

    import struct
    from binascii import unhexlify
    from datetime import datetime, timedelta

    nt_timestamp = struct.unpack("<Q", unhexlify(dt))[0]
    epoch = datetime(1601, 1, 1, 0, 0, 0)
    nt_datetime = epoch + timedelta(microseconds=nt_timestamp / 10)

    return nt_datetime.strftime("%c")

