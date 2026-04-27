import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

from config import NETWORK_PATH_8
from core.ado_client import fetch_builds, download_artifact_zip
from core.deploy_ops import get_next_version, deploy_zip_to_network


class DeployWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Deploy FOXL Build")
        self.geometry("740x560")
        self.resizable(True, True)
        self.minsize(640, 480)
        self._builds = {}
        self._create_widgets()
        self.grab_set()

    def _create_widgets(self):
        # --- Auth ---
        auth_f = ttk.LabelFrame(self, text="Azure DevOps Authentication", padding=(10, 8))
        auth_f.pack(fill="x", padx=10, pady=(8, 4))

        ttk.Label(auth_f, text="Personal Access Token:").pack(side="left")
        self._pat_var = tk.StringVar()
        ttk.Entry(auth_f, textvariable=self._pat_var, show="*", width=48).pack(side="left", padx=6)
        self._load_btn = ttk.Button(auth_f, text="Load Builds", command=self._load_builds)
        self._load_btn.pack(side="left")

        # --- Builds list ---
        builds_f = ttk.LabelFrame(self, text=f"Recent Builds — Pipeline {__import__('config').ADO_PIPELINE_ID}", padding=(10, 8))
        builds_f.pack(fill="both", expand=True, padx=10, pady=4)

        cols = ("build", "result", "finished", "branch", "by")
        self._tree = ttk.Treeview(builds_f, columns=cols, show="headings", selectmode="browse")
        self._tree.heading("build", text="Build #")
        self._tree.heading("result", text="Result")
        self._tree.heading("finished", text="Finished (UTC)")
        self._tree.heading("branch", text="Branch")
        self._tree.heading("by", text="Requested By")
        self._tree.column("build", width=150, minwidth=120)
        self._tree.column("result", width=90, minwidth=80)
        self._tree.column("finished", width=155, minwidth=130)
        self._tree.column("branch", width=120, minwidth=80)
        self._tree.column("by", width=160, minwidth=100)

        sb = ttk.Scrollbar(builds_f, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._tree.bind("<<TreeviewSelect>>", self._on_select)

        # --- Deploy controls ---
        deploy_f = ttk.LabelFrame(self, text="Deployment", padding=(10, 8))
        deploy_f.pack(fill="x", padx=10, pady=4)
        deploy_f.columnconfigure(1, weight=1)

        ttk.Label(deploy_f, text="Target version:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._version_var = tk.StringVar(value="Select a build above")
        ttk.Label(deploy_f, textvariable=self._version_var,
                  font=("TkDefaultFont", 9, "bold")).grid(row=0, column=1, sticky="w")

        self._progress = ttk.Progressbar(deploy_f, length=200, mode="determinate")
        self._progress.grid(row=0, column=2, padx=10, sticky="ew")

        self._deploy_btn = ttk.Button(deploy_f, text="Deploy Selected",
                                      command=self._deploy, state="disabled")
        self._deploy_btn.grid(row=0, column=3, padx=(6, 0))

        # --- Log ---
        log_f = ttk.LabelFrame(self, text="Log", padding=(10, 6))
        log_f.pack(fill="x", padx=10, pady=(4, 8))
        self._log = scrolledtext.ScrolledText(log_f, height=6, wrap=tk.WORD, state="disabled")
        self._log.pack(fill="x")

    # --- Logging ---

    def _log_write(self, msg):
        self._log.configure(state="normal")
        self._log.insert(tk.END, msg + "\n")
        self._log.see(tk.END)
        self._log.configure(state="disabled")

    # --- Load builds ---

    def _load_builds(self):
        pat = self._pat_var.get().strip()
        if not pat:
            messagebox.showwarning("PAT Required", "Enter a Personal Access Token.", parent=self)
            return
        self._load_btn.configure(state="disabled")
        self._deploy_btn.configure(state="disabled")
        self._tree.delete(*self._tree.get_children())
        self._version_var.set("Select a build above")
        self._log_write("Fetching builds from Azure DevOps...")
        threading.Thread(target=self._do_load_builds, args=(pat,), daemon=True).start()

    def _do_load_builds(self, pat):
        ok, result = fetch_builds(pat)
        self.after(0, lambda: self._on_builds_loaded(ok, result))

    def _on_builds_loaded(self, ok, result):
        self._load_btn.configure(state="normal")
        if not ok:
            self._log_write(f"Error loading builds: {result}")
            messagebox.showerror("Load Failed", result, parent=self)
            return

        self._builds = {b["id"]: b for b in result}
        for b in result:
            finished = (b.get("finishTime", "") or "")[:16].replace("T", " ")
            branch = b.get("sourceBranch", "").replace("refs/heads/", "")
            by = b.get("requestedFor", {}).get("displayName", "")
            outcome = b.get("result") or b.get("status", "")
            self._tree.insert("", "end", iid=str(b["id"]),
                              values=(b.get("buildNumber", b["id"]), outcome, finished, branch, by))

        self._log_write(f"Loaded {len(result)} build(s). Select one to deploy.")

    # --- Selection ---

    def _on_select(self, _event):
        if not self._tree.selection():
            self._deploy_btn.configure(state="disabled")
            self._version_var.set("Select a build above")
            return
        self._version_var.set("Checking network share…")
        self._deploy_btn.configure(state="disabled")
        threading.Thread(target=self._resolve_version, daemon=True).start()

    def _resolve_version(self):
        ok, version = get_next_version(NETWORK_PATH_8)
        def update():
            if not self._tree.selection():
                return
            if ok:
                self._version_var.set(version)
                self._deploy_btn.configure(state="normal")
            else:
                self._version_var.set(f"Error: {version}")
                self._deploy_btn.configure(state="disabled")
        self.after(0, update)

    # --- Deploy ---

    def _deploy(self):
        sel = self._tree.selection()
        if not sel:
            return
        build_id = int(sel[0])
        build = self._builds.get(build_id, {})
        build_num = build.get("buildNumber", build_id)
        pat = self._pat_var.get().strip()

        ok, version = get_next_version(NETWORK_PATH_8)
        if not ok:
            messagebox.showerror("Version Error", version, parent=self)
            return

        confirmed = messagebox.askyesno(
            "Confirm Deployment",
            f"Deploy build  {build_num}  as  {version}?\n\n"
            f"Target:\n  {NETWORK_PATH_8}\\{version}\n\n"
            "Steps:\n"
            "  1. Download NetworkShareFiles.zip from Azure DevOps\n"
            "  2. Create the version folder on the network share\n"
            "  3. Extract the zip into that folder",
            parent=self,
        )
        if not confirmed:
            return

        self._deploy_btn.configure(state="disabled")
        self._load_btn.configure(state="disabled")
        self._progress["value"] = 0
        self._log_write(f"─── Deploying build {build_num} as {version} ───")

        threading.Thread(
            target=self._do_deploy, args=(pat, build_id, version), daemon=True
        ).start()

    def _do_deploy(self, pat, build_id, version):
        self.after(0, lambda: self._log_write("Downloading artifact from Azure DevOps..."))

        def on_progress(done, total):
            pct = int(done / total * 100)
            mb_done = done / (1024 * 1024)
            mb_total = total / (1024 * 1024)
            self.after(0, lambda p=pct: self._progress.configure(value=p))
            if pct % 10 == 0:
                msg = f"  {pct}%  ({mb_done:.1f} / {mb_total:.1f} MB)"
                self.after(0, lambda m=msg: self._log_write(m))

        ok, result = download_artifact_zip(pat, build_id, progress_cb=on_progress)
        if not ok:
            self.after(0, lambda: self._finish(False, result))
            return

        zip_path = result
        self.after(0, lambda: self._log_write("Download complete. Extracting to network share..."))
        ok, dest = deploy_zip_to_network(zip_path, version)
        self.after(0, lambda: self._finish(ok, dest))

    def _finish(self, ok, result):
        self._deploy_btn.configure(state="normal")
        self._load_btn.configure(state="normal")
        if ok:
            self._progress["value"] = 100
            self._log_write(f"✓ Done. Deployed to:\n  {result}")
            messagebox.showinfo("Deployment Complete",
                                f"Successfully deployed to:\n{result}", parent=self)
        else:
            self._progress["value"] = 0
            self._log_write(f"✗ Failed: {result}")
            messagebox.showerror("Deployment Failed", result, parent=self)
