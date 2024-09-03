import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
import os
from datetime import datetime

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path

def main():
    root = tk.Tk()
    root.withdraw()

    # 검색어 수집
    search_terms = []
    while True:
        search_term = simpledialog.askstring("입력", "검색할 문구를 입력하세요 (중단하려면 '종료' 입력):")
        if search_term is None or search_term.lower() == '종료':
            break
        if search_term:
            search_terms.append(search_term.lower())

    if not search_terms:
        print("검색할 문구를 입력하지 않았습니다.")
        return

    # 검색어 목록을 윈도우에 표시
    terms_window = tk.Toplevel()
    terms_window.title("입력된 검색어")
    listbox = tk.Listbox(terms_window)
    listbox.pack(fill=tk.BOTH, expand=True)

    for term in search_terms:
        listbox.insert(tk.END, term)

    tk.Button(terms_window, text="확인", command=terms_window.destroy).pack()

    # 파일 선택
    print("A 엑셀 파일을 선택하세요.")
    a_file_path = select_file()
    a_df = pd.read_excel(a_file_path)

    print("B 엑셀 파일을 선택하세요.")
    b_file_path = select_file()
    b_df = pd.read_excel(b_file_path)

    # A/B 파일에서 문구 찾고 필터링
    found_any = False

    for column in a_df.columns:
        a_df[column] = a_df[column].astype(str).str.lower().str.strip()

    for column in b_df.columns:
        b_df[column] = b_df[column].astype(str).str.lower().str.strip()

    for search_term in search_terms:
        found = False
        insert_index = None
        for column in a_df.columns:
            matches = a_df[column].str.contains(search_term, na=False)
            if matches.any():
                found = True
                found_any = True
                insert_index = matches.idxmax() + 1
                print(f"'{search_term}'를 포함하는 데이터가 열 '{column}'에 있습니다.")
                break

        if not found:
            print(f"전체 데이터에서 '{search_term}'를 찾을 수 없습니다.")

        b_filtered_df = pd.DataFrame()
        for column in b_df.columns:
            matches = b_df[column].str.contains(search_term, na=False)
            if matches.any():
                b_filtered_df = pd.concat([b_filtered_df, b_df[matches]], ignore_index=True)

        if not b_filtered_df.empty:
            if insert_index is not None:
                a_df = pd.concat([a_df.iloc[:insert_index], b_filtered_df, a_df.iloc[insert_index:]], ignore_index=True)

    if not found_any:
        print("A 파일에서 검색어들을 찾을 수 없습니다.")
        return

    terms_window.wm_protocol("WM_DELETE_WINDOW", terms_window.destroy)
    root.wait_window(terms_window)

    # 새로운 파일 이름 생성
    base_filename = os.path.splitext(os.path.basename(a_file_path))[0]
    new_file_path = f"{base_filename}_merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    # 새로운 파일로 저장
    a_df.to_excel(new_file_path, index=False)
    print(f"데이터가 성공적으로 결합되어 새로운 파일 '{new_file_path}'로 저장되었습니다.")

if __name__ == "__main__":
    main()
