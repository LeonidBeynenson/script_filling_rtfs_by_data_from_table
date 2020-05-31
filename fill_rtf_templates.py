from pathlib import Path
try:
    from pprint import pprint, pformat
except:
    print("Cannot import pprint")

from copy import deepcopy
import logging as log

ALPHABET_CSV_NAME = "alphabet_for_csv.txt"
ALPHABET_RTF_NAME = "alphabet_for_rtf.rtf"

ALPHABET_CSV_NAME = "alphabet1_forcsv.txt"
ALPHABET_RTF_NAME = "alphabet1_forrtf_changed1.rtf"

RTF_RULANG_PREFIX = r"\hich\f47"
RTF_RULANG_POSTFIX = r"\loch\f47"

DEBUG_MODE = True
if DEBUG_MODE:
    root_logger= log.getLogger()
    root_logger.setLevel(log.DEBUG) # or whatever
    handler = log.FileHandler('log.txt', 'w', 'utf-8') # or whatever
    handler.setFormatter(log.Formatter('%(filename)s:%(lineno)s - %(funcName)s %(message)s')) # or whatever
    root_logger.addHandler(handler)
else:
    root_logger= log.getLogger()
    root_logger.setLevel(log.ERROR) # or whatever
    handler = log.FileHandler('log.txt', 'w', 'utf-8') # or whatever
    handler.setFormatter(log.Formatter('%(filename)s:%(lineno)s - %(funcName)s %(message)s')) # or whatever
    root_logger.addHandler(handler)


def read_alphabets(alphabet_csv_path, alphabet_rtf_path):
    csv_chars = []
    with Path(alphabet_csv_path).open(encoding="utf8") as f_csv:
        for line in f_csv:
            line = line.strip("\n")
            if not line:
                continue
            csv_chars.append(line.strip())

    rtf_charcodes = []
    was_begin = False
    was_first_char = False
    should_stop = False
    with Path(alphabet_rtf_path).open() as f_rtf:
        for line in f_rtf:
            line = line.strip("\n")
            if "[BEGIN]" in line:
                was_begin = True
                continue
            if not was_begin:
                continue
            if line.startswith(r"\par '"):
                was_first_char = True

            if not was_first_char:
                continue

            prefix = r"\par "
            assert line.startswith(prefix), "Line '{}' is not started with prefix in RTF alphabet part".format(line)
            cur_charcode = line[len(prefix):]

            if cur_charcode.endswith("}"):
                cur_charcode = cur_charcode[0:-1]
                should_stop = True

            rtf_charcodes.append(cur_charcode)
            if should_stop:
                break
    if len(csv_chars) != len(rtf_charcodes):
        log.error("Different length of csv_chars and rtf_charcodes!")
        log.error(pformat({"csv_chars": csv_chars}))
        log.error(pformat({"rtf_charcodes": rtf_charcodes}))
        raise RuntimeError("Error in aplhabets")

    a_map = { c: r for c, r in zip(csv_chars, rtf_charcodes) }
    a_map[" "] = " "
    a_map["\n"] = r"\par "

    if DEBUG_MODE:
        log.debug(pformat({"a_map": a_map}))
    return a_map

def split_csv_line(line_iter):
    line = next(line_iter, None)
    if line is None:
        return None
    line = line.strip("\n")
    log.info("begin working with line ='{}'".format(line))

    chunks = []
    cur_part = line
    while cur_part:
        log.debug("CYCLE1: cur_part ='{}'".format(pformat(cur_part)))
        log.debug("        chunks =\n{}".format(pformat(chunks)))
        if not cur_part.startswith('"'):
            ind = cur_part.find(',')
            if ind == -1:
                cur_chunk = cur_part
                cur_part = ""
                chunks.append(cur_chunk)
                continue

            cur_chunk = cur_part[:ind]
            cur_part = cur_part[ind+1:]
            chunks.append(cur_chunk)
            continue

        assert cur_part.startswith('"')
        cur_part = cur_part[1:]
        cur_chunk = ""
        while cur_part:
            # this cycle is for one quoted string,
            # it should start and end with quote char '"'
            # (so, after the last '"' should be either end-of-string or comma)
            # the two chars '""' is used instead of one quote '"'
            log.debug("CYCLE2: cur_part ='{}'".format(pformat(cur_part)))
            log.debug("        cur_chunk ='{}'".format(pformat(cur_chunk)))
            ind = cur_part.find('"')
            if ind == -1:
                # '\n' inside item value
                next_line = next(line_iter, None)
                assert next_line is not None, 'Non-paired chars " in line\n{}\ncur_part=\n{}'.format(line, cur_part)

                next_line = next_line.strip("\n")
                line += "\n" + next_line
                cur_part += "\n" + next_line
                log.debug("        -- getting next line")
                continue

            cur_chunk += cur_part[:ind]
            cur_part = cur_part[ind:]
            if cur_part.startswith('""'):
                cur_chunk += '"'
                cur_part = cur_part[2:]
                continue
            cur_part = cur_part[1:]
            chunks.append(cur_chunk)

            if cur_part and not cur_part.startswith(','):
                log.error("ERROR: cur_part after removing quoted chunk is not started with comma")
                log.error("line  =\n" + line)
                log.error("cur_part =\n" + cur_part)
                raise RuntimeError("ERROR in splitting line")
            if cur_part:
                cur_part = cur_part[1:]
            break

    log.info("end working with line ='{}'".format(line))
    log.info("return chunks =\n{}".format(pformat(chunks)))
    return chunks


def read_csv(csv_path):
    splitted_lines = []
    with Path(csv_path).open(encoding="utf8") as f_csv:
        line_iter = iter(f_csv)
        while True:
            line_chunks = split_csv_line(line_iter)
            if line_chunks is None:
                break
            splitted_lines.append(line_chunks)
    return splitted_lines

def convert_chunk(chunk, a_map):
    assert set(chunk) <= set(a_map.keys())
    conv_chunk = [a_map[c] for c in chunk]
    conv_chunk = "".join(conv_chunk)
    return conv_chunk

def convert_csv(splitted_lines, a_map):
    converted_lines = []
    for line_chunks in splitted_lines:
        cur_conv_line = []
        for chunk in line_chunks:
            if not ( set(chunk) <= set(a_map.keys()) ):
                log.error("ERROR: line is not in alphabet")
                log.error(pformat(line_chunks))
                log.error("chunk =")
                log.error(chunk)
                log.error("absent_chars = {}".format(set(chunk) - set(a_map.keys())))
                a_map = deepcopy(a_map)
                for c in set(chunk) - set(a_map.keys()):
                    a_map[c] = " "
            conv_chunk = convert_chunk(chunk, a_map)
            cur_conv_line.append(conv_chunk)

        converted_lines.append(cur_conv_line)

    return converted_lines

def read_and_convert_csv(csv_path, a_map):
    splitted_lines = read_csv(csv_path)

    column_names = splitted_lines[0]
    column_names = {i:"[" + col_name + "]" for i, col_name in enumerate(column_names) if col_name}
    log.debug(pformat({"column_names": column_names}))


    splitted_lines = splitted_lines[1:]
    converted_lines = convert_csv(splitted_lines, a_map)

    list_map_substs = []
    for conv_line in converted_lines:
        cur_map = {col_name: conv_line[i] for i, col_name in column_names.items()}
        list_map_substs.append(cur_map)

    return list_map_substs


def make_rtf_substitute(src_path, dst_path, map_subst):
    subs_keys = set(map_subst.keys())
    lines = []
    with Path(src_path).open() as f_src:
        for line in f_src:
            new_line = deepcopy(line)
            for k, v in map_subst.items():
                if k in new_line:
                    new_line = new_line.replace(k, RTF_RULANG_PREFIX + " " + v + " " + RTF_RULANG_POSTFIX)
            lines.append(new_line)
    with Path(dst_path).open("w") as f_dst:
        for line in lines:
            f_dst.write(line)

def make_rtf_substitutes_for_all(src_path, dst_prefix, list_map_substs):
    for i, map_subst in enumerate(list_map_substs):
        dst_path = dst_prefix + "_{:04}.rtf".format(i)
        make_rtf_substitute(src_path, dst_path, map_subst)



def main():
    csv_path = "MyCopy_Zapis_NaPraktiky_forchecks__01.csv"
    dst_prefix = "handled"
    src_path = "PREDPISANIE_1.rtf"

    a_map = read_alphabets(ALPHABET_CSV_NAME, ALPHABET_RTF_NAME)
    list_map_substs = read_and_convert_csv(csv_path, a_map)
    make_rtf_substitutes_for_all(src_path, dst_prefix, list_map_substs)

if __name__ == "__main__":
    main()
