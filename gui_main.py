import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import os
import subprocess
import threading
import sys
import datetime

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("自动化报告生成工具")
        self.root.geometry("700x600")
        
        # Style
        style = ttk.Style()
        style.configure("TButton", padding=5)
        style.configure("TLabel", padding=5)

        # Variables
        self.script_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.video_dir_var = tk.StringVar()
        self.video_dir_b_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="普通生成")
        
        # Set default video dir to '视频' in current dir if exists
        default_video = os.path.join(os.getcwd(), "视频")
        if os.path.exists(default_video):
            self.video_dir_var.set(default_video)
        else:
            self.video_dir_var.set(os.getcwd())
        self.video_dir_b_var.set(self.video_dir_var.get())

        # Main Layout
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Configuration Section
        config_frame = ttk.LabelFrame(main_frame, text="配置选项", padding="10")
        config_frame.pack(fill=tk.X, padx=5, pady=5)

        # 1. Script Selection
        ttk.Label(config_frame, text="执行脚本 (.py):").grid(row=0, column=0, sticky=tk.W)
        self.script_combo = ttk.Combobox(config_frame, textvariable=self.script_var, state="readonly", width=50)
        self.script_combo.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(config_frame, text="刷新", command=self.scan_files).grid(row=0, column=2, padx=5)

        # 2. Template Selection
        ttk.Label(config_frame, text="PPT 模板 (.pptx):").grid(row=1, column=0, sticky=tk.W)
        self.template_combo = ttk.Combobox(config_frame, textvariable=self.template_var, state="readonly", width=50)
        self.template_combo.grid(row=1, column=1, padx=5, pady=5)

        # 3. Video Directory Selection
        ttk.Label(config_frame, text="视频文件夹:").grid(row=2, column=0, sticky=tk.W)
        self.video_entry = ttk.Entry(config_frame, textvariable=self.video_dir_var, width=53)
        self.video_entry.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(config_frame, text="浏览...", command=self.browse_video_dir).grid(row=2, column=2, padx=5)

        ttk.Label(config_frame, text="生成模式:").grid(row=3, column=0, sticky=tk.W)
        self.mode_combo = ttk.Combobox(
            config_frame,
            textvariable=self.mode_var,
            state="readonly",
            width=50,
            values=["普通生成", "对比生成(左右分屏)", "已有PPT+新视频对比"],
        )
        self.mode_combo.grid(row=3, column=1, padx=5, pady=5)
        self.mode_combo.bind("<<ComboboxSelected>>", self.on_mode_change)

        self.video_b_label = ttk.Label(config_frame, text="对比视频文件夹(B):")
        self.video_b_entry = ttk.Entry(config_frame, textvariable=self.video_dir_b_var, width=53)
        self.video_b_btn = ttk.Button(config_frame, text="浏览...", command=self.browse_video_dir_b)

        # 4. Existing PPT for Append Mode
        self.existing_ppt_label = ttk.Label(config_frame, text="已有 PPT (.pptx):")
        self.existing_ppt_var = tk.StringVar()
        self.existing_ppt_entry = ttk.Entry(config_frame, textvariable=self.existing_ppt_var, width=53)
        self.existing_ppt_btn = ttk.Button(config_frame, text="浏览...", command=self.browse_existing_ppt)

        # Action Section
        action_frame = ttk.Frame(main_frame, padding="10")
        action_frame.pack(fill=tk.X)
        
        self.run_btn = ttk.Button(action_frame, text="开始生成报告", command=self.start_generation)
        self.run_btn.pack(side=tk.RIGHT, padx=5)
        
        # Author Label
        self.author_label = ttk.Label(action_frame, text="Designed by 草帽", style="Author.TLabel")
        self.author_label.pack(side=tk.RIGHT, padx=20)
        
        self.status_label = ttk.Label(action_frame, text="就绪")
        self.status_label.pack(side=tk.LEFT, padx=5)

        # Log Section
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Initial Scan
        self.scan_files()
        self.on_mode_change()

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def scan_files(self):
        current_dir = os.getcwd()
        py_files = [f for f in os.listdir(current_dir) if f.endswith('.py') and f != os.path.basename(__file__)]
        pptx_files = [f for f in os.listdir(current_dir) if f.endswith('.pptx') and not f.endswith('_generated.pptx') and not f.startswith("~$")]

        self.script_combo['values'] = py_files
        self.template_combo['values'] = pptx_files

        # Set defaults if available
        if 'generate_report.py' in py_files:
            self.script_combo.set('generate_report.py')
        elif py_files:
            self.script_combo.current(0)
            
        if '自动化模板.pptx' in pptx_files:
            self.template_combo.set('自动化模板.pptx')
        elif pptx_files:
            self.template_combo.current(0)

    def browse_video_dir(self):
        directory = filedialog.askdirectory(initialdir=self.video_dir_var.get())
        if directory:
            self.video_dir_var.set(directory)

    def browse_video_dir_b(self):
        directory = filedialog.askdirectory(initialdir=self.video_dir_b_var.get())
        if directory:
            self.video_dir_b_var.set(directory)

    def browse_existing_ppt(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PowerPoint Files", "*.pptx")],
            initialdir=os.getcwd()
        )
        if file_path:
            self.existing_ppt_var.set(file_path)

    def on_mode_change(self, event=None):
        mode = self.mode_var.get()
        
        # Reset UI
        self.video_b_label.grid_remove()
        self.video_b_entry.grid_remove()
        self.video_b_btn.grid_remove()
        self.existing_ppt_label.grid_remove()
        self.existing_ppt_entry.grid_remove()
        self.existing_ppt_btn.grid_remove()

        if mode == "对比生成(左右分屏)":
            self.video_b_label.grid(row=4, column=0, sticky=tk.W)
            self.video_b_entry.grid(row=4, column=1, padx=5, pady=5)
            self.video_b_btn.grid(row=4, column=2, padx=5)
            
            # Auto-select compare script
            if "generate_compare_report.py" in self.script_combo['values']:
                self.script_combo.set("generate_compare_report.py")
                
        elif mode == "已有PPT+新视频对比":
            self.existing_ppt_label.grid(row=4, column=0, sticky=tk.W)
            self.existing_ppt_entry.grid(row=4, column=1, padx=5, pady=5)
            self.existing_ppt_btn.grid(row=4, column=2, padx=5)
            
            # Auto-select append script
            if "generate_append_report.py" in self.script_combo['values']:
                self.script_combo.set("generate_append_report.py")
                
        else: # Normal
            # Auto-select normal script
            if "generate_report.py" in self.script_combo['values']:
                self.script_combo.set("generate_report.py")

    def start_generation(self):
        script = self.script_var.get()
        template = self.template_var.get()
        video_dir = self.video_dir_var.get()
        mode = self.mode_var.get()
        video_dir_b = self.video_dir_b_var.get()
        existing_ppt = self.existing_ppt_var.get()

        if not script or not video_dir:
             messagebox.showerror("错误", "请确保已选择脚本和视频文件夹。")
             return
             
        if mode == "普通生成" and not template:
             messagebox.showerror("错误", "普通生成需要选择 PPT 模板。")
             return

        if mode == "对比生成(左右分屏)" and (not video_dir_b or not template):
            messagebox.showerror("错误", "对比生成需要选择 PPT 模板和对比视频文件夹(B)。")
            return
            
        if mode == "已有PPT+新视频对比" and not existing_ppt:
            messagebox.showerror("错误", "请选择已有的 PPT 文件。")
            return

        self.run_btn.config(state='disabled')
        self.status_label.config(text="正在运行...")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        thread = threading.Thread(target=self.run_subprocess, args=(script, template, video_dir, mode, video_dir_b, existing_ppt))
        thread.daemon = True
        thread.start()

    def run_subprocess(self, script, template, video_dir, mode, video_dir_b, existing_ppt):
        try:
            # Generate output filename based on video directory and date
            # Format: {VideoFolderName}-模流分析报告-{YYYY.MM.DD}.pptx
            
            # Normalize path to strip trailing slashes for basename to work
            video_folder_name = os.path.basename(os.path.normpath(video_dir))
            current_date = datetime.datetime.now().strftime("%Y.%m.%d")
            
            # Absolute paths
            cwd = os.getcwd()
            
            if mode == "对比生成(左右分屏)":
                output_path = f"{video_folder_name}-模流分析报告-{current_date}.pptx"
                compare_script = os.path.join(cwd, "generate_compare_report.py")
                if not os.path.exists(compare_script):
                    raise FileNotFoundError("缺少对比脚本 generate_compare_report.py")
                script_path = compare_script
                
                template_path = os.path.join(cwd, template)
                output_full_path = os.path.join(cwd, output_path)
                
                cmd = [
                    sys.executable,
                    script_path,
                    "--video_dir_a",
                    video_dir,
                    "--video_dir_b",
                    video_dir_b,
                    "--template_path",
                    template_path,
                    "--output_path",
                    output_full_path,
                ]
            elif mode == "已有PPT+新视频对比":
                # Output name logic for append mode
                # Format: {VideoFolderName}-第二模-模流分析报告-{YYYY.MM.DD}.pptx
                output_path = f"{video_folder_name}-第二模-模流分析报告-{current_date}.pptx"
                
                append_script = os.path.join(cwd, "generate_append_report.py")
                if not os.path.exists(append_script):
                    raise FileNotFoundError("缺少脚本 generate_append_report.py")
                script_path = append_script
                
                output_full_path = os.path.join(cwd, output_path)
                
                cmd = [
                    sys.executable,
                    script_path,
                    "--existing_ppt",
                    existing_ppt,
                    "--video_dir",
                    video_dir,
                    "--output_path",
                    output_full_path,
                ]
                
            else: # Normal
                output_path = f"{video_folder_name}-模流分析报告-{current_date}.pptx"
                script_path = os.path.join(cwd, script)
                template_path = os.path.join(cwd, template)
                output_full_path = os.path.join(cwd, output_path)
                
                cmd = [
                    sys.executable,
                    script_path,
                    "--video_dir",
                    video_dir,
                    "--template_path",
                    template_path,
                    "--output_path",
                    output_full_path,
                ]
            
            self.log(f"执行命令: {' '.join(cmd)}")
            
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                encoding='gbk', # Use gbk for Windows console
                errors='replace',
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )

            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                if line:
                    self.root.after(0, self.log, line.strip())

            return_code = process.poll()
            if return_code == 0:
                self.root.after(0, lambda: self.status_label.config(text="完成"))
                self.root.after(0, lambda: messagebox.showinfo("成功", f"报告已生成：\n{output_path}"))
            else:
                self.root.after(0, lambda: self.status_label.config(text="出错"))
                self.root.after(0, lambda: messagebox.showerror("错误", "生成过程中发生错误，请查看日志。"))

        except Exception as e:
            self.root.after(0, self.log, f"系统错误: {str(e)}")
            self.root.after(0, lambda: self.status_label.config(text="系统错误"))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state='normal'))

if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()
