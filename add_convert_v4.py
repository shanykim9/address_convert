import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import requests
import re
import os
import threading
import sys
import time

# ==========================================
# [설정 영역]
# ==========================================
API_KEY = "devU01TX0FVVEgyMDI2MDEyMDExNTIxNjExNzQ2MTE=" 
ADDRESS_COL = "주소"

# 디자인 테마 설정 (화이트 모드)
ctk.set_appearance_mode("Light")  
ctk.set_default_color_theme("blue") 
# ==========================================

class AddressConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 1. 윈도우 설정
        self.title("Address Converter Pro (Final Edition)")
        self.geometry("1100x750") 
        self.configure(fg_color="#F5F6FA") 

        # 변수 초기화
        self.file_list = []
        self.output_dir = os.getcwd()

        # 2. UI 레이아웃 구성
        self.create_widgets()

    def create_widgets(self):
        # 전체 컨테이너
        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.pack(fill="both", expand=True, padx=20, pady=20)

        # ---------------------------------------------------------
        # [왼쪽 패널] 컨트롤러
        # ---------------------------------------------------------
        self.left_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        self.left_frame.pack(side="left", fill="both", expand=True, padx=(0, 20))

        # [헤더] 타이틀
        self.lbl_title = ctk.CTkLabel(self.left_frame, text="주소 변환 시스템", 
                                      font=("Pretendard", 28, "bold"), text_color="#2D3436")
        self.lbl_title.pack(anchor="w", pady=(0, 5))
        
        self.lbl_subtitle = ctk.CTkLabel(self.left_frame, text="Dashboard Controller", 
                                      font=("Pretendard", 12), text_color="#636E72")
        self.lbl_subtitle.pack(anchor="w", pady=(0, 20))

        # [섹션 1] 파일 선택 카드
        self.card_files = ctk.CTkFrame(self.left_frame, fg_color="white", border_color="#DFE6E9", border_width=1, corner_radius=10)
        self.card_files.pack(fill="x", pady=(0, 15))

        self.lbl_files = ctk.CTkLabel(self.card_files, text="1. 변환할 파일 선택", 
                                      font=("Pretendard", 16, "bold"), text_color="#2D3436")
        self.lbl_files.pack(pady=15, padx=20, anchor="w")

        # 버튼 그룹
        self.btn_frame = ctk.CTkFrame(self.card_files, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=20, pady=(0, 10))

        self.btn_add_file = ctk.CTkButton(self.btn_frame, text="+ 파일 추가", command=self.add_files, width=100, height=32,
                                          fg_color="#0984E3", hover_color="#74B9FF", text_color="white")
        self.btn_add_file.pack(side="left", padx=(0, 10))
        
        self.btn_add_folder = ctk.CTkButton(self.btn_frame, text="+ 폴더 선택", command=self.add_folder, width=100, height=32,
                                            fg_color="transparent", border_width=1, border_color="#B2BEC3", 
                                            text_color="#636E72", hover_color="#F1F2F6")
        self.btn_add_folder.pack(side="left", padx=(0, 10))

        self.btn_clear = ctk.CTkButton(self.btn_frame, text="초기화", command=self.clear_list, width=80, height=32,
                                       fg_color="transparent", text_color="#D63031", hover_color="#FFEAA7")
        self.btn_clear.pack(side="right")

        # 파일 목록창
        self.txt_file_list = ctk.CTkTextbox(self.card_files, height=120, state="disabled", 
                                            fg_color="#F8F9FA", text_color="#2D3436", font=("Pretendard", 12))
        self.txt_file_list.pack(fill="x", padx=20, pady=(0, 20))

        # [섹션 2] 저장 경로 카드
        self.card_output = ctk.CTkFrame(self.left_frame, fg_color="white", border_color="#DFE6E9", border_width=1, corner_radius=10)
        self.card_output.pack(fill="x", pady=(0, 15))

        self.lbl_output_title = ctk.CTkLabel(self.card_output, text="2. 결과 저장 위치", 
                                             font=("Pretendard", 16, "bold"), text_color="#2D3436")
        self.lbl_output_title.pack(pady=15, padx=20, anchor="w")

        self.output_frame = ctk.CTkFrame(self.card_output, fg_color="transparent")
        self.output_frame.pack(fill="x", padx=20, pady=(0, 20))

        self.entry_output = ctk.CTkEntry(self.output_frame, placeholder_text=self.output_dir, 
                                         fg_color="#F8F9FA", border_color="#DFE6E9", text_color="#2D3436")
        self.entry_output.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.entry_output.insert(0, self.output_dir)
        self.entry_output.configure(state="disabled")

        self.btn_output = ctk.CTkButton(self.output_frame, text="변경", command=self.change_output_dir, width=60,
                                        fg_color="#636E72", hover_color="#2D3436", text_color="white")
        self.btn_output.pack(side="right")

        # [섹션 3] 진행 상황 카드
        self.card_progress = ctk.CTkFrame(self.left_frame, fg_color="white", border_color="#DFE6E9", border_width=1, corner_radius=10)
        self.card_progress.pack(fill="x", pady=(0, 15))

        self.lbl_progress_title = ctk.CTkLabel(self.card_progress, text="3. 실시간 진행 상황", 
                                               font=("Pretendard", 16, "bold"), text_color="#2D3436")
        self.lbl_progress_title.pack(pady=15, padx=20, anchor="w")

        self.progress_bar = ctk.CTkProgressBar(self.card_progress, progress_color="#00B894", height=12)
        self.progress_bar.pack(fill="x", padx=20, pady=(0, 8))
        self.progress_bar.set(0)

        self.lbl_status = ctk.CTkLabel(self.card_progress, text="작업 대기 중...", text_color="#636E72", font=("Pretendard", 12))
        self.lbl_status.pack(pady=(0, 20))

        # [실행 버튼]
        self.btn_run = ctk.CTkButton(self.left_frame, text="변환 시작 (START)", command=self.start_thread,
                                     height=55, font=("Pretendard", 18, "bold"), corner_radius=10,
                                     fg_color="#2D3436", hover_color="#000000")
        self.btn_run.pack(fill="x", pady=(0, 20), side="bottom")

        # ---------------------------------------------------------
        # [오른쪽 패널] 상세 로그
        # ---------------------------------------------------------
        self.right_frame = ctk.CTkFrame(self.container, fg_color="white", border_color="#DFE6E9", border_width=1, corner_radius=10)
        self.right_frame.pack(side="right", fill="both", expand=True)

        self.lbl_log = ctk.CTkLabel(self.right_frame, text="상세 로그 (System Log)", 
                                    font=("Pretendard", 16, "bold"), text_color="#2D3436")
        self.lbl_log.pack(pady=15, padx=20, anchor="w")

        # 글자 크기 18로 유지
        self.txt_log = ctk.CTkTextbox(self.right_frame, font=("Consolas", 18), 
                                      fg_color="#F8F9FA", border_color="#DFE6E9", border_width=1, text_color="#2D3436")
        self.txt_log.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.txt_log.configure(state="disabled")


    # --- 기능 함수들 ---
    def log(self, msg):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def update_file_list_display(self):
        self.txt_file_list.configure(state="normal")
        self.txt_file_list.delete("1.0", "end")
        if not self.file_list:
            self.txt_file_list.insert("end", "선택된 파일이 없습니다.")
        else:
            for f in self.file_list:
                self.txt_file_list.insert("end", f"📄 {os.path.basename(f)}\n")
        self.txt_file_list.configure(state="disabled")

    def add_files(self):
        files = filedialog.askopenfilenames(title="엑셀 파일 선택", filetypes=[("Excel files", "*.xlsx;*.xls")])
        count = 0
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
                count += 1
        self.update_file_list_display()
        if count > 0: self.log(f"파일 {count}개가 추가되었습니다.")

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
                            count += 1
            self.update_file_list_display()
            self.log(f"폴더 스캔 완료: {count}개 파일 추가됨.")

    def clear_list(self):
        self.file_list = []
        self.update_file_list_display()
        self.log("목록이 초기화되었습니다.")

    def change_output_dir(self):
        folder = filedialog.askdirectory(title="결과를 저장할 폴더 선택")
        if folder:
            self.output_dir = folder
            self.entry_output.configure(state="normal")
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, self.output_dir)
            self.entry_output.configure(state="disabled")
            self.log(f"저장 위치 변경됨: {folder}")

    def start_thread(self):
        if not self.file_list:
            messagebox.showwarning("경고", "변환할 파일을 먼저 추가해주세요!")
            return
        
        self.btn_run.configure(state="disabled", text="데이터 처리 중...", fg_color="#636E72")
        t = threading.Thread(target=self.run_conversion)
        t.start()

    # --- 변환 엔진 ---
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
        self.log("="*40)
        self.log(">>> [시스템 시작] 변환 작업을 시작합니다.")
        
        total_files = len(self.file_list)
        
        for idx, file_path in enumerate(self.file_list):
            file_name = os.path.basename(file_path)
            self.log(f"\n[{idx+1}/{total_files}] 분석 중: {file_name}")
            
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
                total_rows = len(df)
                
                self.progress_bar.set(0)
                
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
                    
                    current_count = i + 1
                    progress_val = current_count / total_rows
                    self.progress_bar.set(progress_val)
                    
                    percent = progress_val * 100
                    status_msg = f"처리 중... {current_count}/{total_rows} ({percent:.1f}%) | 성공: {len(success_list)}건"
                    self.lbl_status.configure(text=status_msg, text_color="#0984E3")
                    
                    if current_count % 5 == 0:
                        self.update_idletasks()

                base_name = os.path.splitext(file_name)[0]
                
                # 저장 로그 출력
                if success_list:
                    save_name = f"{base_name}_변환성공.xlsx"
                    save_path = os.path.join(self.output_dir, save_name)
                    pd.DataFrame(success_list).to_excel(save_path, index=False)
                    self.log(f"   👍 저장 완료 (성공 {len(success_list)}건)")
                
                if fail_list:
                    save_name = f"{base_name}_변환실패.xlsx"
                    save_path = os.path.join(self.output_dir, save_name)
                    pd.DataFrame(fail_list).to_excel(save_path, index=False)
                    self.log(f"   ⚠️ 저장 완료 (실패 {len(fail_list)}건)")
                
                # [추가] 변환 성공률 계산 및 표시
                if total_rows > 0:
                    success_rate = (len(success_list) / total_rows) * 100
                    self.log(f"   📊 변환 성공률: {success_rate:.1f}%")

            except Exception as e:
                self.log(f"❌ 오류 발생: {e}")

        self.progress_bar.set(1.0)
        self.lbl_status.configure(text="모든 작업 완료!", text_color="#00B894")
        self.log("\n>>> [종료] 모든 변환 작업이 끝났습니다.")
        messagebox.showinfo("완료", "작업이 성공적으로 끝났습니다.")
        self.btn_run.configure(state="normal", text="변환 시작 (START)", fg_color="#2D3436")

if __name__ == "__main__":
    app = AddressConverterApp()
    app.mainloop()