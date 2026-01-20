import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk  # [추가] 막대 그래프(Progress Bar)용 모듈
import pandas as pd
import requests
import re
import os
import threading
import sys

# ==========================================
# [설정 영역]
# ==========================================
API_KEY = "devU01TX0FVVEgyMDI2MDEyMDExNTIxNjExNzQ2MTE=" 
ADDRESS_COL = "주소"
# ==========================================

class AddressConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("주소 변환기 Pro v2.0 (시각화 패치)")
        self.root.geometry("600x850") # 높이를 조금 늘림
        
        # 변수 초기화
        self.file_list = []
        self.output_dir = os.getcwd()

        # UI 구성
        self.create_widgets()

    def create_widgets(self):
        # 1. 파일 선택 영역
        frame_top = tk.LabelFrame(self.root, text="1. 변환할 파일 선택", padx=10, pady=10)
        frame_top.pack(fill="x", padx=10, pady=5)

        btn_add_file = tk.Button(frame_top, text="파일 추가", command=self.add_files, width=15)
        btn_add_file.grid(row=0, column=0, padx=5, pady=5)
        
        btn_add_folder = tk.Button(frame_top, text="폴더 선택", command=self.add_folder, width=15)
        btn_add_folder.grid(row=0, column=1, padx=5, pady=5)
        
        btn_clear = tk.Button(frame_top, text="목록 초기화", command=self.clear_list, width=15, fg="red")
        btn_clear.grid(row=0, column=2, padx=5, pady=5)

        self.listbox = tk.Listbox(frame_top, height=5, selectmode="extended")
        self.listbox.grid(row=1, column=0, columnspan=3, sticky="ew", pady=5)

        # 2. 저장 경로 영역
        frame_mid = tk.LabelFrame(self.root, text="2. 결과 저장 위치", padx=10, pady=10)
        frame_mid.pack(fill="x", padx=10, pady=5)

        self.lbl_output = tk.Label(frame_mid, text=self.output_dir, fg="blue", wraplength=450)
        self.lbl_output.pack(side="left", fill="x", expand=True)
        
        btn_output = tk.Button(frame_mid, text="폴더 변경", command=self.change_output_dir)
        btn_output.pack(side="right")

        # 3. [신규] 진행률 표시 영역 (그래프바)
        frame_progress = tk.LabelFrame(self.root, text="3. 진행 상황 (Real-time)", padx=10, pady=10)
        frame_progress.pack(fill="x", padx=10, pady=5)

        # 막대 그래프 (Progress Bar)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=self.progress_var)
        self.progress_bar.pack(fill="x", expand=True, pady=5)

        # 숫자 표시 라벨 (예: 50/100 (50%))
        self.lbl_progress = tk.Label(frame_progress, text="대기 중...", font=("Arial", 10, "bold"), fg="green")
        self.lbl_progress.pack()

        # 4. 실행 버튼
        self.btn_run = tk.Button(self.root, text="변환 시작 (START)", command=self.start_thread, 
                                 bg="darkblue", fg="white", font=("Arial", 12, "bold"), height=2)
        self.btn_run.pack(fill="x", padx=10, pady=10)

        # 5. 상세 로그 창
        frame_log = tk.LabelFrame(self.root, text="상세 로그", padx=10, pady=10)
        frame_log.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_area = scrolledtext.ScrolledText(frame_log, state='disabled', height=8)
        self.log_area.pack(fill="both", expand=True)

    # --- 기능 함수들 ---
    def log(self, msg):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    def add_files(self):
        files = filedialog.askopenfilenames(title="엑셀 파일 선택", filetypes=[("Excel files", "*.xlsx;*.xls")])
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
                self.listbox.insert(tk.END, f)
        self.log(f"파일 {len(files)}개 추가됨.")

    def add_folder(self):
        folder = filedialog.askdirectory(title="폴더 선택")
        if folder:
            count = 0
            for root_dir, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith(('.xlsx', '.xls')) and "변환성공" not in file and "변환실패" not in file:
                        full_path = os.path.join(root_dir, file)
                        if full_path not in self.file_list:
                            self.file_list.append(full_path)
                            self.listbox.insert(tk.END, full_path)
                            count += 1
            self.log(f"폴더에서 엑셀 파일 {count}개 발견하여 추가함.")

    def clear_list(self):
        self.file_list = []
        self.listbox.delete(0, tk.END)
        self.log("목록이 초기화되었습니다.")

    def change_output_dir(self):
        folder = filedialog.askdirectory(title="결과를 저장할 폴더 선택")
        if folder:
            self.output_dir = folder
            self.lbl_output.config(text=folder)
            self.log(f"저장 위치 변경됨: {folder}")

    def start_thread(self):
        if not self.file_list:
            messagebox.showwarning("경고", "변환할 파일을 먼저 추가해주세요!")
            return
        
        self.btn_run.config(state="disabled", text="변환 중입니다...", bg="gray")
        t = threading.Thread(target=self.run_conversion)
        t.start()

    # --- 핵심 변환 로직 ---
    def advanced_clean_text(self, text):
        if not isinstance(text, str): return str(text)
        text = text.replace('.', '')
        text = re.sub(r'\(.*?\)', ' ', text)
        text = text.replace('~', ' ')
        text = re.sub(r'(동|리|가|읍|면)([0-9가-힣])', r'\1 \2', text)
        text = re.sub(r'(번지)([0-9가-힣])', r'\1 \2', text)
        text = re.sub(r'(\d)([가-힣])', r'\1 \2', text)
        text = re.sub(r'(\d) (동|호|층|번지)', r'\1\2', text)
        text = re.sub(r'([가-힣])(\d)', r'\1 \2', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    def get_road_address(self, keyword):
        url = "https://business.juso.go.kr/addrlink/addrLinkApi.do"
        params = {
            "confmKey": API_KEY, "currentPage": 1, "countPerPage": 1,
            "keyword": keyword, "resultType": "json"
        }
        try:
            response = requests.get(url, params=params, timeout=5)
            if response.status_code == 200:
                data = response.json()
                if data['results']['common']['totalCount'] != '0':
                    return data['results']['juso'][0]
        except:
            pass
        return None

    def run_conversion(self):
        self.log("="*30)
        self.log(">>> 변환 작업을 시작합니다.")
        
        for idx, file_path in enumerate(self.file_list):
            file_name = os.path.basename(file_path)
            self.log(f"\n[{idx+1}/{len(self.file_list)}] 파일 처리 중: {file_name}")
            
            try:
                df = None
                for engine in [None, 'openpyxl', 'xlrd']:
                    try:
                        df = pd.read_excel(file_path, engine=engine) if engine else pd.read_excel(file_path)
                        break
                    except: continue
                
                if df is None:
                    self.log(f"❌ 읽기 실패: {file_name}")
                    continue
                
                if ADDRESS_COL not in df.columns:
                    self.log(f"⚠️ 실패: '{ADDRESS_COL}' 컬럼 없음")
                    continue

                success_list = []
                fail_list = []
                total = len(df)
                
                # [그래프 초기화] 해당 파일의 전체 건수로 설정
                self.progress_bar['maximum'] = total
                self.progress_var.set(0)
                
                for i, row in df.iterrows():
                    original_addr = str(row[ADDRESS_COL])
                    cleaned = self.advanced_clean_text(original_addr)
                    
                    words = cleaned.split()
                    result = None
                    current_words = words[:]
                    
                    while len(current_words) >= 2:
                        res = self.get_road_address(" ".join(current_words))
                        if res:
                            result = res
                            break
                        current_words.pop()
                    
                    if result:
                        r_dict = row.to_dict()
                        r_dict['도로명주소'] = result['roadAddr']
                        r_dict['우편번호'] = result['zipNo']
                        success_list.append(r_dict)
                    else:
                        fail_list.append(row.to_dict())
                    
                    # [그래프 업데이트] 매 건마다 갱신
                    current_count = i + 1
                    self.progress_var.set(current_count)
                    progress_percent = (current_count / total) * 100
                    
                    # 라벨 업데이트 (예: 처리 중: 50 / 100 (50.0%) - 성공 45건)
                    status_text = f"처리 중: {current_count} / {total} ({progress_percent:.1f}%)  [성공: {len(success_list)}건]"
                    self.lbl_progress.config(text=status_text)
                    
                    # UI 강제 새로고침 (부드러운 움직임을 위해)
                    # 너무 자주하면 느려질 수 있으니 5건마다 한 번씩만 UI 렌더링
                    if current_count % 5 == 0:
                        self.root.update_idletasks()

                # 파일 저장
                base_name = os.path.splitext(file_name)[0]
                if success_list:
                    save_name = f"{base_name}_변환성공.xlsx"
                    save_path = os.path.join(self.output_dir, save_name)
                    pd.DataFrame(success_list).to_excel(save_path, index=False)
                    self.log(f"   👍 성공 저장 완료 ({len(success_list)}건)")
                
                if fail_list:
                    save_name = f"{base_name}_변환실패.xlsx"
                    save_path = os.path.join(self.output_dir, save_name)
                    pd.DataFrame(fail_list).to_excel(save_path, index=False)
                    self.log(f"   ⚠️ 실패 저장 완료 ({len(fail_list)}건)")

            except Exception as e:
                self.log(f"❌ 에러 발생: {e}")

        # 완료 처리
        self.progress_var.set(0) # 그래프 초기화
        self.lbl_progress.config(text="모든 작업 완료! (대기 중)")
        self.log("\n모든 작업이 완료되었습니다!")
        messagebox.showinfo("완료", "모든 변환 작업이 끝났습니다.")
        self.btn_run.config(state="normal", text="변환 시작 (START)", bg="darkblue")

if __name__ == "__main__":
    root = tk.Tk()
    app = AddressConverterApp(root)
    root.mainloop()