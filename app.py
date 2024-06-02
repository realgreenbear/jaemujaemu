from flask import Flask, request, render_template, redirect, url_for
from werkzeug.utils import secure_filename
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def find_value_in_sheet(filepath, sheet_names, keywords):
    try:
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(filepath, sheet_name=sheet_name)
            except Exception as e:
                continue

            for keyword in keywords:
                sheet_df = df[df.eq(keyword).any(axis=1)]

                if sheet_df is None:
                    continue

                if not sheet_df.empty:
                    value = sheet_df.iloc[0, 1]
                    return value

        # 모든 시트에서 값을 찾지 못한 경우
        return None
        
    except Exception as e:
        print("Error occurred while finding value in sheet:", e)
        return None
    
def calculate_percentage(a, b):
    try:
        if b == 0 or pd.isna(a) or pd.isna(b):
            return None
        result = (a / b) * 100
        return round(result, 2)
    except Exception as e:
        print("Error occurred while calculating percentage:", e)
        return None
    
def duldul(a, b):
    if a is not None:
        return a
    else:
        return b

def calculate_ratios(filepath):
    try:
        # 매출총이익률
        a11 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["매출총이익"])
        a12 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["매출총이익(손실)"])
        a1 = duldul(a11, a12)
        a2 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["매출액", "수익(매출액)"])
        if a1 is None or a2 is None:
            a3 = 100
        else:
            if a1 == a11:
                a3 = calculate_percentage(a1, a2)
            else:
                a3 = calculate_percentage(a1, a2) * (-1)

        # 영업이익률
        b11 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["영업이익"])
        b12 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["영업이익(손실)"])
        b1 = duldul(b11, b12)
        b2 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["매출액", "수익(매출액)", "영업수익"])
        if b1 is None or b2 is None:
            b3 = 100
        else:
            if b1 == b11:
                b3 = calculate_percentage(b1, b2)
            else:
                b3 = calculate_percentage(b1, b2) * (-1)

        # 순이익률
        c11 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["당기순이익", "분기순이익"])
        c12 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["당기순이익(손실)","분기순이익(손실)"])
        c1 = duldul(c11, c12)
        c2 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["매출액", "수익(매출액)", "영업수익"])
        if c1 is None or c2 is None:
            c3 = 100
        else:
            if c1 == c11:
                c3 = calculate_percentage(c1, c2)
            else:
                c3 = calculate_percentage(c1, c2) * (-1)

        # 자산수익률
        d11 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["당기순이익", "분기순이익"])
        d12 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["당기순이익(손실)","분기순이익(손실)"])
        d1 = duldul(d11, d12)
        d2 = find_value_in_sheet(filepath, ["D210000"], ["    자산총계"])
        if d1 is None or d2 is None:
            d3 = 100
        else:
            if d1 == d11:
                d3 = calculate_percentage(d1, d2)
            else:
                d3 = calculate_percentage(d1, d2) * (-1)

        # 자기자본수익률
        e11 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["당기순이익", "분기순이익"])
        e12 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["당기순이익(손실)","분기순이익(손실)"])
        e1 = duldul(e11, e12)
        e2 = find_value_in_sheet(filepath, ["D210000"], ["    자본총계"])
        if e1 is None or e2 is None:
            e3 = 100
        else:
            if e1 == e11:
                e3 = calculate_percentage(e1, e2)
            else:
                e3 = calculate_percentage(e1, e2) * (-1)

        # 유동비율
        f1 = find_value_in_sheet(filepath, ["D210000"], ["    유동자산"])
        f2 = find_value_in_sheet(filepath, ["D210000"], ["    유동부채"])
        f3 = calculate_percentage(f1, f2)

        # 당좌비율
        g11 = find_value_in_sheet(filepath, ["D210000"], ["    유동자산"])
        g12 = find_value_in_sheet(filepath, ["D210000"], ["        재고자산", "        유동재고자산"])
        g1 = g11 - g12
        g2 = find_value_in_sheet(filepath, ["D210000"], ["    유동부채"])
        if g1 is None or g2 is None:
            g3 = 100
        else:
            g3 = calculate_percentage(g1, g2)

        # 현금비율
        h1 = find_value_in_sheet(filepath, ["D210000"], ["        현금및현금성자산", "        현금 및 현금성자산"])
        h2 = find_value_in_sheet(filepath, ["D210000"], ["    유동부채"])
        if h1 is None or h2 is None:
            h3 = 100
        else:
            h3 = calculate_percentage(h1, h2)

        # 부채비율
        i1 = find_value_in_sheet(filepath, ["D210000"], ["    부채총계"])
        i2 = find_value_in_sheet(filepath, ["D210000"], ["    자본총계"])
        i3 = calculate_percentage(i1, i2)

        # 자기자본비율
        j1 = find_value_in_sheet(filepath, ["D210000"], ["    자본총계"])
        j2 = find_value_in_sheet(filepath, ["D210000"], ["    자산총계"])
        if j1 is None or j2 is None:
            j3 = 100
        else:
            j3 = calculate_percentage(j1, j2)

        # 이자보상비율
        k1 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["금융비용", "금융원가"])
        k21 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["영업이익"])
        k22 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["영업이익(손실)"])
        k2 = duldul(k21, k22)
        if k1 is None or k2 is None:
            k3 = 100
        else:
            if k2 == k21:
                k3 = calculate_percentage(k2, k1)
            else:
                k3 = calculate_percentage(k2, k1) * (-1)
        # 재고자산회전율
        l1 = find_value_in_sheet(filepath, ["D210000"], ["        재고자산", "        유동재고자산"])
        l2 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["매출원가"])
        if l1 is None or l2 is None:
            l3 = 0
        else:
            l3 = calculate_percentage(l2, l1)

        # 매출채권회전율
        m1 = find_value_in_sheet(filepath, ["D210000"], ["        매출채권", "        매출채권 및 기타채권", "        매출채권및미수금"])
        m2 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["매출액", "수익(매출액)", "영업수익"])
        if m1 is None or m2 is None:
            m3 = 100
        else:
            m3 = calculate_percentage(m2, m1)

        # 총자산회전율
        n1 = find_value_in_sheet(filepath, ["D210000"], ["    자산총계"])
        n2 = find_value_in_sheet(filepath, ["D310000", "D431410", "D431420"], ["매출액", "수익(매출액)", "영업수익"])
        if n1 is None or n2 is None:
            n3 = 100
        else:
            n3 = calculate_percentage(n2, n1)

        return a3, b3, c3, d3, e3, f3, g3, h3, i3, j3, k3, l3, m3, n3
    except Exception as e:
        print("Error occurred while calculating ratios:", e)
        return None, None, None, None, None, None, None, None, None, None, None, None, None, None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            a3, b3, c3, d3, e3, f3, g3, h3, i3, j3, k3, l3, m3, n3 = calculate_ratios(filepath)
            return f"매출총이익률: {a3}%<br>영업이익율: {b3}%<br>순이익률: {c3}%<br>자산수익률: {d3}%<br>자기자본수익률: {e3}%<br>유동비율: {f3}%<br>당좌비율: {g3}%<br>현금비율: {h3}%<br>부채비율: {i3}%<br>자기자본비율: {j3}%<br>이자보상비율: {k3}%<br>재고자산회전율: {l3}%<br>매출채권회전율: {m3}%<br>총자산회전율: {n3}%"

if __name__ == '__main__':
    app.run(debug=True)
