import glob
import mojimoji



def main():
    # ファイルの指定
    # file_name = "test_output.txt"

    for file_name in glob.glob(r'.\*\*.txt'):
        # 既存ファイルの読み込み
        with open(file_name, encoding="utf-8") as f:
            textData = f.read()
        
        # テキストのキレイキレイ
        textData = clean_text(textData)    

        # ファイルの保存
        with open(file_name, mode="w", encoding="utf-8") as f:
            f.write(textData)


# テキストの置換内容
transtable = str.maketrans({
    ' ':'', 
    '\t':'',    
    '、':'',
    '。':'',
    '！':'',
    '？':'',
    })

def clean_text(target_text):
    target_text = target_text.translate(transtable)
    target_text = mojimoji.han_to_zen(target_text)
    return target_text

if __name__ == "__main__":
    main()



'''
Refereneces URL

Pythonで空白等を削除する方法 = https://note.com/skilltunejp/n/n1fbf37edd494
半角と全角を変換する ＝ https://note.nkmk.me/python-str-convert-full-half-width/#mojimojihan_to_zen 
既存ファイルの書き換え ＝ https://gammasoft.jp/blog/text-file-edit-by-python/
複数ファイルの一括編集 ＝ https://qiita.com/wataame1011/questions/877f3cc9109b2abf0c81
'''