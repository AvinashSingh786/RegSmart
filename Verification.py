import hashlib
import os

original = []
copy = []


def verify_dump(path):
    global copy
    global original
    tmp = load_hash_file(path)
    if tmp == "No hash files found":
        return 2
    if tmp:
        # print(original)
        # print(copy)
        if verify_dumps_from_original(path):
            print("\n\n++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
            print("\t\tVerified")
            print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
            return True
        else:
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            print("\t\tError - Original file and forensic copy integrity not verified")
            return False
    else:
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print("\t\tError - Hash files do not match please verify")
    return False


def hash_file(path):
    md5 = hashlib.md5()
    with open(path, 'rb') as file:
        md5.update(file.read())

    return md5.hexdigest()


def get_hash(path):
    try:
        return hash_file(path)
    except Exception:
        return get_dir_hash(path)


def get_dir_hash(directory):
    try:
        hashes = []
        for path, dirs, files in os.walk(directory):
            for file in sorted(files):  # we sort to guarantee that files will always go in the same order
                hashes.append(hash_file(os.path.join(path, file)))
            for dir in sorted(dirs):  # we sort to guarantee that dirs will always go in the same order
                hashes.append(get_dir_hash(os.path.join(path, dir)))
            break  # we only need one iteration - to get files and dirs in current directory
        md5 = hashlib.md5()
        md5.update(''.join(hashes).encode())
        return md5.hexdigest()
    except Exception as ee:
        print(ee)


def verify_dumps_from_original(path):
    k = 0
    print("Verifying ...")
    print("-----------------------------------------------------------------------------------------------------------")
    for filename in os.listdir(path):
        if filename.endswith(".regacquire") or filename == "DEFAULT" or filename == "NTUSER.DAT" or filename == "SAM" \
                or filename == "SECURITY" or filename == "SOFTWARE" or filename == "SYSTEM":
            a = get_hash(path+"/"+filename)
            b = get_hash(path+"/original/"+filename)
            if filename == "HKEY_USERS.regacquire":
                print(filename + "\t\t\t\t" + a + "\t" + b)
            else:
                print(filename + "\t\t" + a + "\t" + b)
            if a == b:
                if original[k] == hash_file(path+"/"+filename):
                    continue
            else:
                return False
            k += 1
    return True


def load_hash_file(name):
    global original
    try:
        with open(name+"/original/hash.hash", 'r', encoding='utf-8') as f:
            for line in f:
                original.append(line.strip().split(" ")[3])

        global copy
        with open(name+"/hash.hash", 'r', encoding='utf-8') as f:
            for line in f:
                copy.append(line.strip().split(" ")[3])

        return original == copy
    except FileNotFoundError:
        return "No hash files found"


def is_valid_regacquire(name):
    try:
        with open(name + "/original/hash.hash", 'r', encoding='utf-8') as f:
            return True

    except FileNotFoundError:
        return False


# verify_dump(
#     "F:/UP_2017_CS_HN_S1/COS700/Projects/DF-Research/Program/Analysis/RegSmart/Avinash_DESKTOP-IAP0BLH_2017-06-27")
