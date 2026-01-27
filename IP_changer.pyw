import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
from pathlib import Path
import os
import sys
import ctypes

# --- 管理者権限チェック ---
if not ctypes.windll.shell32.IsUserAnAdmin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 0)
    sys.exit()

# --- BASE_DIR の修正（exe 対応） ---
if getattr(sys, 'frozen', False):
    # exe 化された場合は実行ファイルのフォルダ
    BASE_DIR = Path(sys.executable).parent
else:
    # 普通の Python 実行
    BASE_DIR = Path(__file__).parent

CONFIG_FILE = BASE_DIR / "config.json"
IP_PRESETS_FILE = BASE_DIR / "ip_presets.json"
NO_WINDOW = 0x08000000

# --- ロジック関数 ---
def get_wired_interfaces():
    ps = "Get-NetAdapter | Where-Object { $_.PhysicalMediaType -eq '802.3' } | Select-Object -ExpandProperty Name"
    r = subprocess.run(["powershell", "-Command", ps], capture_output=True, text=True, creationflags=NO_WINDOW)
    return [l for l in r.stdout.splitlines() if l.strip()] or []

def prefix_to_netmask(prefix_str):
    try:
        prefix = int(prefix_str)
        mask = (0xffffffff >> (32 - prefix)) << (32 - prefix)
        return ".".join(str((mask >> i) & 0xff) for i in [24, 16, 8, 0])
    except:
        return "255.255.255.0"

# --- メインアプリ ---
class UltimateIPTool:
    def __init__(self, root):
        self.root = root
        self.root.title("IP一発切替 Pro")
        self.root.geometry("450x650")
        self.presets = self.load_presets()
        self.create_widgets()
        self.root.after(100, self.initial_setup)

    def load_presets(self):
        if IP_PRESETS_FILE.exists():
            try:
                with open(IP_PRESETS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                pass
        return {}

    def initial_setup(self):
        interfaces = get_wired_interfaces()
        self.combo_if["values"] = interfaces
        last_iface = None
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    last_iface = json.load(f).get("last_interface")
            except:
                pass

        if last_iface in interfaces:
            self.iface_var.set(last_iface)
        elif interfaces:
            self.iface_var.set(interfaces[0])

        self.update_current_info()

    def create_widgets(self):
        main = ttk.Frame(self.root, padding="15")
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text="対象アダプター:").pack(anchor=tk.W)
        self.iface_var = tk.StringVar()
        self.combo_if = ttk.Combobox(main, textvariable=self.iface_var, state="readonly")
        self.combo_if.pack(fill=tk.X, pady=5)
        self.combo_if.bind("<<ComboboxSelected>>", lambda e: self.update_current_info())

        ttk.Separator(main, orient='horizontal').pack(fill='x', pady=10)

        ttk.Label(main, text="IPプリセット適用:", font=("Meiryo", 10, "bold")).pack(anchor=tk.W)
        self.btn_frame = ttk.Frame(main)
        self.btn_frame.pack(fill=tk.X, pady=5)
        self.refresh_preset_buttons()

        ttk.Separator(main, orient='horizontal').pack(fill='x', pady=10)

        info_frame = ttk.LabelFrame(main, text=" 現在のIPステータス ", padding="10")
        info_frame.pack(fill=tk.X, pady=10)
        self.info_text = tk.Text(info_frame, height=5, font=("Consolas", 10), bg="#f0f0f0", relief="flat")
        self.info_text.pack(fill=tk.X)

        btn_grid = ttk.Frame(main)
        btn_grid.pack(fill=tk.X, pady=5)
        ttk.Button(btn_grid, text="🔄 再読込", command=self.update_current_info).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_grid, text="🌐 設定を開く", command=lambda: subprocess.Popen("ncpa.cpl", shell=True)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        self.lbl_loading = ttk.Label(main, text="", foreground="blue", font=("Meiryo", 9))
        self.lbl_loading.pack(pady=5)

        ttk.Separator(main, orient='horizontal').pack(fill='x', pady=10)

        ttk.Button(main, text="✚ プリセットを新規追加", command=self.add_preset_dialog).pack(fill=tk.X, pady=10)

    def refresh_preset_buttons(self):
        for w in self.btn_frame.winfo_children():
            w.destroy()
        for name in self.presets.keys():
            ttk.Button(self.btn_frame, text=f"▶ {name} を適用",
                       command=lambda n=name: self.apply_ip(n)).pack(fill=tk.X, pady=2)

    def add_preset_dialog(self):
        name = simpledialog.askstring("新規追加", "設定名を入力してください:")
        if not name: return
        ip = simpledialog.askstring("新規追加", f"『{name}』のIPアドレス:")
        if not ip: return
        gw = simpledialog.askstring("新規追加", f"『{name}』のゲートウェイ:")
        if not gw: return
        self.presets[name] = {"ip": ip, "mask": "255.255.255.0", "gw": gw}
        try:
            with open(IP_PRESETS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.presets, f, indent=4, ensure_ascii=False)
            self.refresh_preset_buttons()
            messagebox.showinfo("成功", f"プリセット『{name}』を追加しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"保存失敗: {e}")

    def update_current_info(self):
        iface = self.iface_var.get()
        if not iface: return

        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"last_interface": iface}, f, ensure_ascii=False)

        self.lbl_loading.config(text="読み込みに数秒かかります・・")
        self.root.update_idletasks()

        try:
            ps = f'Get-NetIPConfiguration -InterfaceAlias "{iface}" | ForEach-Object {{ "$($_.IPv4Address.IPAddress)|$($_.IPv4Address.PrefixLength)|$($_.IPv4DefaultGateway.NextHop)" }}'
            r = subprocess.run(["powershell", "-Command", ps], capture_output=True, text=True, creationflags=NO_WINDOW)
            self.info_text.delete("1.0", tk.END)
            if "|" in r.stdout:
                ipv4, prefix, gw = r.stdout.strip().split("|")
                self.info_text.insert(tk.END, f"IP Address  : {ipv4}\nSubnet Mask : {prefix_to_netmask(prefix)}\nGateway     : {gw}")
            else:
                self.info_text.insert(tk.END, "設定なし")
        except:
            self.info_text.insert(tk.END, "取得エラー")
        finally:
            self.lbl_loading.config(text="")

    def apply_ip(self, name):
        data = self.presets[name]
        iface = self.iface_var.get()
        cmd = f'netsh interface ipv4 set address name="{iface}" source=static addr={data["ip"]} mask={data["mask"]} gateway={data["gw"]}'
        try:
            subprocess.run(cmd, shell=True, creationflags=NO_WINDOW)
            messagebox.showinfo("成功", f"{name} を適用しました")
            self.update_current_info()
        except:
            messagebox.showerror("エラー", "適用失敗")

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateIPTool(root)
    root.mainloop()
