import win32com.client as win32
from pathlib import Path
import logging as log
import sys

root_logger= log.getLogger()
root_logger.setLevel(log.DEBUG) # or whatever
handler = log.FileHandler('log_conversion.txt', 'w', 'utf-8') # or whatever
handler.setLevel(log.ERROR)
handler.setFormatter(log.Formatter('%(filename)s:%(lineno)s - %(funcName)s %(message)s')) # or whatever
root_logger.addHandler(handler)

streamhandler = log.StreamHandler(sys.stdout)
streamhandler.setLevel(log.DEBUG)
streamhandler.setFormatter(log.Formatter('%(filename)s: %(message)s')) # or whatever
root_logger.addHandler(streamhandler)

PDF_FILETYPE_CODE = 17

def main():
    print("Close all Word documents!")
    enter = input("Press Enter when all is closed")

    cur_path = Path(__file__).parent.resolve()
    docs_paths = list(cur_path.glob("*.rtf"))
    log.info("Found {} rtf files".format(len(docs_paths)))

    log.info("Open Word")
    word = win32.gencache.EnsureDispatch('Word.Application')
    assert word, "Cannot open Word"
    # word.Visible = False
    log.info("Word is opened successfully")

    num_successes = 0
    N = len(docs_paths)
    for i, p in enumerate(docs_paths):
        prefix = "#{}/{}: ".format(i+1, N)
        log.info(prefix + "Working with {}".format(p.name))
        try:
            doc = word.Documents.Open(str(p.resolve()))
        except:
            log.error(prefix + "Error: cannot open {}".format(p.name))
            log.error(prefix + "       The full path {}".format(p))
            log.error(prefix)
            continue
        if not doc:
            log.error(prefix + "Strange error: Cannot open {}".format(p.name))
            log.error(prefix + "               The full path {}".format(p))
            log.error(prefix)
            continue

        log.info(prefix + "The file {} is opened".format(p.name))
        p2 = p.with_suffix(".pdf")
        try:
            doc.SaveAs(str(p2), 17)
        except:
            log.error(prefix + "ERROR: The file {} cannot be saved as pdf".format(p.name))
            log.error(prefix + "       The full path {}".format(p))
            log.error(prefix + "       The full target path {}".format(p2))
            log.error(prefix)
            doc.Close()
            continue

        log.info(prefix + "The file {} is saved as {}".format(p.name, p2.name))
        num_successes += 1
        doc.Close()

    log.error("Everything is done -- {}/{} files converted successfully".format(num_successes, len(docs_paths)))

if __name__ == "__main__":
    main()
