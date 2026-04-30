# GUI.py
# ver 1.0.1

import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import os
import threading
import queue # 안전한 스레드를 위해 큐 사용(get은 값을 가져오는 함수 / put은 큐에 데이터를 집어넣는 함수)
import merge_logic  # 분리한 엑셀 병합 로직 파일

# 전역 변수
MAX_FILES = 20 # 선택할 수 있는 파일의 최대 개수
selected_files = [] # 현재 선택된 파일들의 경로(문자열 리스트)를 저장하는 리스트
gui_queue = queue.Queue() # 스레드 간 안전한 통신을 위한 큐 생성

"""
GUI 프로그램에서 blocking이 왜 치명적인가

1. 큐가 비어 있음
2. get()에서 멈춤
3. Tkinter 이벤트 루프가 멈춤

결과: 창이 얼어붙음 (응답 없음)
"""

def process_queue(): 
    # 큐에 있는 모든 메시지를 처리하여 GUI를 안전하게 업데이트
    try:
        while True: # 큐가 비어있을 때까지 루프를 돌며 쌓인 메시지를 모두 처리
            """
                message = gui_queue.get() get함수는 기본적으로 블로킹 함수 이 코드의 의미는
                >>> 큐에 뭔가 들어올 때까지 이 줄에서 아무것도 하지 말고 기다려라

                message = gui_queue.get_nowait()
                이 줄의 정확한 의미
                >>> “큐에 지금 당장 꺼낼 게 있으면 가져오고
                >>> 없으면 기다리지 말고
                >>> 대신 예외를 던져라”
            """
            message = gui_queue.get_nowait() # 블로킹을 방지하기 위해서는 get_nowait 함수를 사용하면 됨 / 블로킹 없이 메시지를 가져옴
            command, data = message

            if command == 'log':
                log_message(data)
            elif command == 'show_error':
                title, msg = data
                messagebox.showerror(title, msg) 
            elif command == 'show_info':
                title, msg = data
                messagebox.showinfo(title, msg)
            elif command == 'task_done':
                merge_button.config(state=tk.NORMAL)
                select_button.config(state=tk.NORMAL)
                delete_button.config(state=tk.NORMAL)
                up_button.config(state=tk.NORMAL)
                down_button.config(state=tk.NORMAL)

    except queue.Empty: # 큐가 비어있으면 get_nowait()이 이 예외를 발생시킴 / 처리할 메시지가 더 이상 없다는 의미이므로 루프 종료
        pass
        """
        실제 동작은 이렇다

        * 큐에 메시지가 있을 때
        ('log', '파일 처리 중')
            → 정상적으로 message에 저장됨

        * 큐가 비어 있을 때
        queue.Empty  # 예외 발생

        except queue.Empty: # 그래서 이 코드 필요
            pass
        """
    finally:
        root.after(100, process_queue) # 모든 메시지를 처리한 후, 100ms 뒤에 다시 큐를 확인할 것을 예약합니다.

def select_files(): # 파일 선택 대화상자를 열고(askopenfilenames 함수) 선택된 파일들을 리스트박스에 표시
    global selected_files
    
    files = filedialog.askopenfilenames( # 여러 파일을 선택할 수 있는 파일 선택 대화상자 열기(튜플로 반환)
        title="병합할 엑셀 파일을 선택하세요",
        filetypes=[("Excel 파일", "*.xlsx")] # openpyxl 라이브러리는 엑셀 구버전인 xls 파일과 호환되지 않음
    )
    
    if files: # 만약 사용자가 파일을 하나라도 선택했다면? (취소 버튼을 누르지 않았다면)
        new_files_count = 0 # 사용자가 선택한 파일 중에서, 실제로 목록에 추가 성공한 파일의 개수를 세는 역할

        for file in files:
            if file not in selected_files: # 새로 선택된 파일 중, 기존 목록에 없는 파일만 추가
                if len(selected_files) < MAX_FILES: # 현재 선택된 파일들의 개수가 20개 미만인가?
                    selected_files.append(file)
                    new_files_count += 1
                else:
                    log_message(f"오류: 파일은 최대 {MAX_FILES}개까지 선택할 수 있습니다.")
                    messagebox.showwarning(
                        "선택 개수 초과",
                        f"파일은 최대 {MAX_FILES}개까지만 선택할 수 있습니다.\n\n"
                        f"일부 파일은 추가되지 않았습니다."
                    )
                    break 
        
        if new_files_count > 0:
            log_message(f"{new_files_count}개의 파일이 새로 추가되었습니다. (총 {len(selected_files)}개)")
            log_message("-" * 20)

            update_file_listbox()

def delete_selected_file(event=None): # 리스트박스에서 선택된 파일을 목록에서 삭제
    global selected_files
    selection_indices = file_listbox.curselection() # curselection()은 선택된 항목의 인덱스를 튜플 형태로 반환 (예: (2,))
    
    if not selection_indices:
        messagebox.showwarning("알림", "삭제할 파일을 목록에서 선택해주세요.")
        return
        
    selected_index = selection_indices[0]
    
    deleted_file_path = selected_files.pop(selected_index)
    file_name = os.path.basename(deleted_file_path)
    
    log_message(f"'{file_name}' 파일이 목록에서 삭제되었습니다.")
    update_file_listbox()

def move_file_up(): # 선택한 파일을 목록에서 한 칸 위로 이동
    global selected_files
    selection_indices = file_listbox.curselection() # curselection()은 선택된 항목의 인덱스를 튜플 형태로 반환 (예: (2,))

    if not selection_indices:
        messagebox.showwarning("알림", "순서를 변경할 파일을 목록에서 선택해주세요.")
        return
        
    selected_index = selection_indices[0]
    
    if selected_index > 0:
        new_index = selected_index - 1
        selected_files[selected_index], selected_files[new_index] = selected_files[new_index], selected_files[selected_index]
        
        file_name = os.path.basename(selected_files[new_index])
        log_message(f"'{file_name}' 순서가 변경되었습니다.")
        
        update_file_listbox(select_index=new_index)

def move_file_down(): # 선택한 파일을 목록에서 한 칸 아래로 이동
    global selected_files
    selection_indices = file_listbox.curselection() # curselection()은 선택된 항목의 인덱스를 튜플 형태로 반환 (예: (2,))

    if not selection_indices:
        messagebox.showwarning("알림", "순서를 변경할 파일을 목록에서 선택해주세요.")
        return
        
    selected_index = selection_indices[0]
    
    if selected_index < len(selected_files) - 1:
        new_index = selected_index + 1
        selected_files[selected_index], selected_files[new_index] = selected_files[new_index], selected_files[selected_index]
        
        file_name = os.path.basename(selected_files[new_index])
        log_message(f"'{file_name}' 순서가 변경되었습니다.")
        
        update_file_listbox(select_index=new_index) 

def update_file_listbox(select_index=None): 
   # tk.END는 Listbox에서 항목들 중 가장 마지막 항목을 가리키는 tkinter 상수
    file_listbox.delete(0, tk.END) # 리스트박스의 모든 내용을 지움 / delete(start_idx, end_idx) start_idx 부터 end_idx 까지 항목 삭제(end_idx 포함)

    # --- 플레이스홀더 제어 로직 ---
    if not selected_files:
        listbox_placeholder.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
    else:
        listbox_placeholder.place_forget() # 목록이 있으면 숨김
    # ----------------------------

    log_message(f"현재 {len(selected_files)}개의 파일이 목록에 있습니다. (병합 순서)")

    for i, file_path in enumerate(selected_files): # selected_files 리스트의 내용을 다시 채움 (번호 포함)
        file_name = os.path.basename(file_path)
        file_listbox.insert(tk.END, f"{i+1}. {file_name}") # 리스트박스의 맨 끝에 "새 항목"을 추가
        log_message(f"  {i+1}. {file_name}")
    log_message("-" * 20)

    if select_index is not None and select_index < file_listbox.size(): # 만약 select_index가 주어졌다면, 해당 항목을 선택 상태로 만듦
        """
        file_listbox.selection_set(2)       # 인덱스 2번 항목을 선택
        file_listbox.selection_set(2, 5)    # 인덱스 2~5 항목을 모두 선택
        """
        file_listbox.selection_set(select_index) # 지정한 인덱스 범위를 선택 상태로 만듦 / 선택된 항목은 리스트박스에서 파란색(또는 강조 색)으로 표시됨
        file_listbox.activate(select_index) # 해당 인덱스를 활성 항목(active item)으로 설정/ 활성 항목은 사용자가 키보드로 ↑↓ 이동할 때 기준이 되는 위치
        file_listbox.see(select_index) # 지정한 인덱스가 리스트박스 화면에 보이도록 스크롤을 자동 조정(항목이 이미 보이는 위치에 있다면 아무 변화 없음)

def log_message(message): # 로그 영역에 메시지를 추가하는 함수
    log_area.config(state='normal') # (state='normal') >>> 편집 가능 상태를 의미(사용자가 마우스/키보드로 내용을 자유롭게 수정 가능)
    log_area.insert(tk.END, message + "\n") # 메시지를 입력하는 함수 / insert(입력 위치, 입력할 내용)
    log_area.config(state='disabled') # (state='disabled') >>> 편집 불가 상태를 의미(읽기 전용)
    log_area.see(tk.END) # 맨 마지막 텍스트가 보이도록 스크롤을 마지막 텍스트가 출력된 위치로 이동
    root.update_idletasks() # 위젯의 크기나 위치 변경 등 화면에 표시되어야 할 변경 사항을 즉시 반영하도록 함

def start_merge_thread(): # 메인 병합 함수
    # 파일 저장 경로를 먼저 묻고, 정해지면 스레드를 시작

    # 작업 시작 전 파일 목록 유효성 검사
    if not selected_files: # 파일이 하나도 선택되지 않았다면
        messagebox.showerror("오류", "병합할 엑셀 파일을 먼저 선택해주세요.")
        return
    elif len(selected_files) < 2: # 2개 미만이면 병합할 의미가 없으므로 오류와 함께 종료
        messagebox.showerror("오류", "병합하려면 최소 2개 이상의 파일을 선택해야 합니다.")
        return
    
    """
        asksaveasfilename() 메서드는 파일을 저장할 수 있는 대화 상자를 엽니다. 이 메서드는 사용자가 지정한 파일 이름을 반환함
        만약 취소 버튼을 누르면 빈 문자열을 반환합니다. 이 메서드를 사용하면 사용자가 원하는 파일 이름으로 쉽게 저장할 수 있음
    """
    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel 통합 문서", "*.xlsx")],
        title="병합된 파일 저장 위치 선택"
    )

    # 사용자가 경로를 선택한 경우에만 스레드 시작
    if not output_path: # 저장 대화상자에서 취소 버튼을 눌렀다면
        log_message("저장이 취소되었습니다.")
        return

    # <<<--- 수정된 부분 시작: 원본 파일 덮어쓰기 방지 기능 추가 --->>>
    normalized_output_path = os.path.normpath(output_path) # os.path.normpath를 사용하여 경로 구분자( / 또는 \ )를 통일시켜 정확하게 비교(경로 정규화)
    normalized_input_paths = [os.path.normpath(f) for f in selected_files]

    if normalized_output_path in normalized_input_paths:
        log_message("오류: 저장하려는 파일 이름이 선택된 파일 중 하나와 동일합니다.")
        messagebox.showerror(
            "덮어쓰기 오류",
            "원본 파일을 덮어쓸 수 없습니다.\n\n"
            "병합된 파일은 다른 이름으로 저장해주세요."
        )
        return
    # <<<--- 수정된 부분 끝 --->>>

    merge_button.config(state=tk.DISABLED) # 병합 시작 및 저장 버튼 비활성화
    select_button.config(state=tk.DISABLED) # 병합할 파일 추가 버튼 비활성화
    delete_button.config(state=tk.DISABLED) # 선택 파일 삭제 버튼 비활성화
    up_button.config(state=tk.DISABLED) # 위로 버튼 비활성화
    down_button.config(state=tk.DISABLED) # 아래로 버튼 비활성화
    
    # 스레드에 merge_logic 파일의 함수를 연결하고, 인자를 전달
    merge_thread = threading.Thread(
        target=merge_logic.merge_excel_files, # target에 실행할 함수를 지정
        args=(output_path, selected_files, gui_queue) # target에 지정한 함수의 매개 변수에 전달할 값 지정
    )
    # 데몬 스레드는 메인 스레드가 종료되면 같이 종료되는 스레드 즉, 프로그램 종료 시 병합 프로세스가 깔끔하게 끝나도록 설정
    merge_thread.daemon = True # True로 설정하면 메인 프로그램(프로세스)이 종료될 때 이 스레드는 자동으로 강제 종료 
    merge_thread.start() 

def GUI():
    global root, file_listbox, log_area, merge_button, select_button, delete_button, up_button, down_button
    global listbox_placeholder # 플레이스홀더 제어를 위해 전역 변수 추가

    # 기본 창 설정
    root = tk.Tk()
    root.title("Excel 파일 병합 프로그램")
    root.geometry("700x550")
    
    # --- 상단 프레임: 파일 선택 ---
    top_frame = tk.Frame(root, bd=2, relief=tk.GROOVE)
    top_frame.pack(padx=10, pady=10, fill=tk.X)
    
    button_frame = tk.Frame(top_frame)
    button_frame.pack(side=tk.LEFT, padx=5, pady=5)

    # 이제 버튼들의 부모(master)를 top_frame이 아닌 button_frame으로 지정합니다.
    # pack() 옵션에서 side=tk.TOP을 사용하여 버튼들을 위에서 아래로 쌓습니다.
    select_button = tk.Button(button_frame, text="병합할 파일 추가", width=15, command=select_files)
    select_button.pack(side=tk.TOP, fill=tk.X, pady=(0, 5)) # fill=tk.X는 버튼 너비를 프레임에 맞추는 기능, pady는 버튼 사이 간격

    delete_button = tk.Button(button_frame, text="선택 파일 삭제", width=15, command=delete_selected_file)
    delete_button.pack(side=tk.TOP, fill=tk.X)
    
    # 선택된 파일 목록을 보여줄 리스트박스
    file_listbox = tk.Listbox(top_frame, height=5, bg="white")
    file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5), pady=5)

    # --- 플레이스홀더 추가 시작 ---
    listbox_placeholder = tk.Label(
        file_listbox, # 리스트박스 내부에 배치
        text="비밀번호가 걸린 파일은 병합이 정상적으로 진행되지 않습니다\n비밀번호가 걸려있는 경우 비밀번호를 해제한 후 병합을 진행해 주세요",
        fg="gray",
        bg="white"
    )
    # 리스트박스 정중앙에 배치
    listbox_placeholder.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
    # --- 플레이스홀더 추가 끝 ---

    # 사용자가 리스트박스에서 항목을 선택하고 키보드의 'Delete' 키를 누르면 delete_selected_file 함수가 호출됨
    file_listbox.bind("<Delete>", delete_selected_file) 
    
    # 순서 변경 버튼 프레임 및 버튼 생성
    reorder_frame = tk.Frame(top_frame)
    reorder_frame.pack(side=tk.LEFT, padx=(0, 5), pady=5)

    # 업 버튼
    up_button = tk.Button(reorder_frame, text="▲ 위로", width=8, command=move_file_up)
    up_button.pack(side=tk.TOP, padx=5, pady=(0, 5))

    # 다운 버튼
    down_button = tk.Button(reorder_frame, text="▼ 아래로", width=8, command=move_file_down)
    down_button.pack(side=tk.TOP, padx=5)

    # 안내 메시지 프레임
    instruction_frame = tk.Frame(root)
    instruction_frame.pack(padx=10, pady=(0, 5), fill=tk.X)

    instruction_text = f"※ 안내: 병합할 엑셀 파일을 최대 {MAX_FILES}개까지 추가할 수 있으며, 목록의 순서대로 병합됩니다."
    instruction_label = tk.Label(instruction_frame, text=instruction_text, fg="blue", anchor='w', justify=tk.LEFT)
    instruction_label.pack(fill=tk.X)
    
    # 주의 메시지 프레임
    instruction_frame2 = tk.Frame(root)
    instruction_frame2.pack(padx=10, pady=(0, 5), fill=tk.X)

    instruction_text2 = f"※ 주의: 병합할 엑셀 파일의 개수가 많거나 엑셀 파일 내의 내용이 많다면 소요 시간이 길어질 수 있습니다."
    instruction_label2 = tk.Label(instruction_frame2, text=instruction_text2, fg="red", anchor='w', justify=tk.LEFT)
    instruction_label2.pack(fill=tk.X)

    # --- 중간 프레임: 병합 실행 ---
    middle_frame = tk.Frame(root)
    middle_frame.pack(padx=10, pady=5, fill=tk.X)
    
    merge_button = tk.Button(middle_frame, text="병합 시작 및 저장", font=("Arial", 10, "bold"), command=start_merge_thread)
    merge_button.pack(fill=tk.X, ipady=5)

    # --- 하단 프레임: 로그 표시 ---
    bottom_frame = tk.Frame(root, bd=2, relief=tk.GROOVE)
    bottom_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    log_label = tk.Label(bottom_frame, text="진행 과정 로그")
    log_label.pack(anchor='w', padx=5)

    # 로그를 보여줄 스크롤 가능한 텍스트 영역
    log_area = scrolledtext.ScrolledText(bottom_frame, wrap=tk.WORD, bg="white", state='disabled')
    log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    root.after(100, process_queue) # GUI가 시작된 후 큐 확인 프로세스를 시작
    
    root.mainloop()

if __name__ == "__main__":
    GUI()