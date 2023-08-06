# ヒトリ
from googletrans import Translator, constants
from pprint import pprint

def main():
    while True:
        ts = Translator()
        i_text = str(input("input text :"))
        out = ts.translate(i_text)
        if out.src == "en":
            out = ts.translate(i_text,dest="ja")
            print(f"translated: {out.text}")
        else:
            out = ts.translate(i_text,src="ja",dest="en")
            print(f"translated: {out.text}")
        
if __name__ == '__main__':
    main()
