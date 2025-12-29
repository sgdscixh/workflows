if __name__=="__main__":
    from PyPDF4 import PdfFileMerger
    import sys, os
    merger = PdfFileMerger(strict=False)
    print(sys.argv)
    # base = os.path.dirname(sys.argv[1])
    # for p in sys.argv[1:]:
    #     merger.append(fileobj=open(p, 'rb'), import_bookmarks=True)
    # merger.write(fileobj=open(f'merge.pdf', 'wb'))
