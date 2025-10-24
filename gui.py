"""
Trial Orders Automation - GUI Application
Professional interface for managing automated court document processing
Supports: Claude, OpenAI, Gemini, and vLLM (self-hosted)
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import queue
import json
import os
from pathlib import Path
from datetime import datetime
import sys

# Import the automation module
from automation import TrialOrdersAutomation, Config


class AutomationGUI:
    """Main GUI Application"""

    def __init__(self, root):
        self.root = root
        self.root.title("Trial Orders Automation - Multi-Provider AI")
        self.root.geometry("1000x750")
        self.root.minsize(900, 650)

        # Queue for thread-safe logging
        self.log_queue = queue.Queue()

        # Automation instance
        self.automation = None
        self.is_running = False
        self.process_thread = None

        # Configuration file
        self.config_file = Path("config.json")

        # Setup UI
        self.setup_ui()

        # Load saved configuration
        self.load_config()

        # Start log queue checker
        self.check_log_queue()

    def setup_ui(self):
        """Setup the user interface"""

        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Tab 1: Main Control
        self.tab_main = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_main, text="  Main  ")
        self.setup_main_tab()

        # Tab 2: Configuration
        self.tab_config = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_config, text="  Configuration  ")
        self.setup_config_tab()

        # Tab 3: About
        self.tab_about = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_about, text="  About  ")
        self.setup_about_tab()

    def setup_main_tab(self):
        """Setup main control tab"""

        # Status Frame
        status_frame = ttk.LabelFrame(self.tab_main, text="Status", padding=10)
        status_frame.pack(fill=tk.X, padx=10, pady=5)

        self.status_label = ttk.Label(
            status_frame,
            text="Ready to process emails",
            font=("Segoe UI", 10)
        )
        self.status_label.pack()

        self.progress = ttk.Progressbar(
            status_frame,
            mode='indeterminate',
            length=300
        )
        self.progress.pack(pady=5)

        # Control Frame
        control_frame = ttk.Frame(self.tab_main)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        self.start_button = ttk.Button(
            control_frame,
            text="▶ Start Processing",
            command=self.start_processing,
            width=20
        )
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = ttk.Button(
            control_frame,
            text="⏹ Stop",
            command=self.stop_processing,
            state=tk.DISABLED,
            width=20
        )
        self.stop_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(
            control_frame,
            text="Clear Log",
            command=self.clear_log,
            width=15
        ).pack(side=tk.RIGHT, padx=5)

        # Stats Frame
        stats_frame = ttk.LabelFrame(self.tab_main, text="Statistics", padding=10)
        stats_frame.pack(fill=tk.X, padx=10, pady=5)

        stats_grid = ttk.Frame(stats_frame)
        stats_grid.pack()

        ttk.Label(stats_grid, text="Emails Processed:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.emails_processed_label = ttk.Label(stats_grid, text="0", font=("Segoe UI", 10, "bold"))
        self.emails_processed_label.grid(row=0, column=1, sticky=tk.W)

        ttk.Label(stats_grid, text="PDFs Downloaded:").grid(row=0, column=2, sticky=tk.W, padx=20)
        self.pdfs_downloaded_label = ttk.Label(stats_grid, text="0", font=("Segoe UI", 10, "bold"))
        self.pdfs_downloaded_label.grid(row=0, column=3, sticky=tk.W)

        ttk.Label(stats_grid, text="Last Run:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.last_run_label = ttk.Label(stats_grid, text="Never")
        self.last_run_label.grid(row=1, column=1, sticky=tk.W, pady=5)

        # Log Frame
        log_frame = ttk.LabelFrame(self.tab_main, text="Activity Log", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create scrolled text widget for logs
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            width=80,
            height=20,
            font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Configure tags for colored output
        self.log_text.tag_config("INFO", foreground="black")
        self.log_text.tag_config("SUCCESS", foreground="green")
        self.log_text.tag_config("WARNING", foreground="orange")
        self.log_text.tag_config("ERROR", foreground="red")

    def setup_config_tab(self):
        """Setup configuration tab"""

        # Create canvas with scrollbar
        canvas = tk.Canvas(self.tab_config)
        scrollbar = ttk.Scrollbar(self.tab_config, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Email Configuration (Outlook COM)
        email_frame = ttk.LabelFrame(scrollable_frame, text="Email Configuration (Outlook COM)", padding=10)
        email_frame.pack(fill=tk.X, padx=10, pady=5)

        info_label = ttk.Label(
            email_frame,
            text="✅ Uses your existing Outlook installation - no configuration needed!\n\n"
                 "Make sure Outlook is installed and logged in on this computer.",
            foreground="green",
            font=("Segoe UI", 9)
        )
        info_label.pack(pady=10)

        # OneDrive/SharePoint Configuration
        sp_frame = ttk.LabelFrame(scrollable_frame, text="OneDrive/SharePoint (Local Sync)", padding=10)
        sp_frame.pack(fill=tk.X, padx=10, pady=5)

        info_label = ttk.Label(
            sp_frame,
            text="✅ Uses local OneDrive folders - no credentials needed!\n\n"
                 "Files sync to SharePoint automatically via OneDrive.\n"
                 "Make sure OneDrive is syncing properly.",
            foreground="green",
            font=("Segoe UI", 9)
        )
        info_label.pack(pady=10)

        # AI Provider Configuration
        ai_frame = ttk.LabelFrame(scrollable_frame, text="AI Provider Configuration", padding=10)
        ai_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(ai_frame, text="AI Provider:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.ai_provider_var = tk.StringVar(value="claude")
        provider_combo = ttk.Combobox(
            ai_frame,
            textvariable=self.ai_provider_var,
            values=["claude", "openai", "gemini", "vllm"],
            state="readonly",
            width=47
        )
        provider_combo.grid(row=0, column=1, padx=5, pady=5)
        provider_combo.bind("<<ComboboxSelected>>", self.on_provider_changed)

        # Claude Configuration
        self.claude_frame = ttk.Frame(ai_frame)
        self.claude_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W+tk.E, pady=5)

        ttk.Label(self.claude_frame, text="Anthropic API Key:").pack(side=tk.LEFT)
        self.anthropic_key_entry = ttk.Entry(self.claude_frame, width=50, show="*")
        self.anthropic_key_entry.pack(side=tk.LEFT, padx=5)

        # OpenAI Configuration
        self.openai_frame = ttk.Frame(ai_frame)
        self.openai_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W+tk.E, pady=5)

        ttk.Label(self.openai_frame, text="OpenAI API Key:").pack(side=tk.LEFT)
        self.openai_key_entry = ttk.Entry(self.openai_frame, width=50, show="*")
        self.openai_key_entry.pack(side=tk.LEFT, padx=5)

        ttk.Label(self.openai_frame, text="Model:").pack(side=tk.LEFT, padx=(20, 0))
        self.openai_model_entry = ttk.Entry(self.openai_frame, width=20)
        self.openai_model_entry.pack(side=tk.LEFT, padx=5)
        self.openai_model_entry.insert(0, "gpt-4o")

        # Gemini Configuration
        self.gemini_frame = ttk.Frame(ai_frame)
        self.gemini_frame.grid(row=3, column=0, columnspan=2, sticky=tk.W+tk.E, pady=5)

        ttk.Label(self.gemini_frame, text="Gemini API Key:").pack(side=tk.LEFT)
        self.gemini_key_entry = ttk.Entry(self.gemini_frame, width=50, show="*")
        self.gemini_key_entry.pack(side=tk.LEFT, padx=5)

        ttk.Label(self.gemini_frame, text="Model:").pack(side=tk.LEFT, padx=(20, 0))
        self.gemini_model_entry = ttk.Entry(self.gemini_frame, width=20)
        self.gemini_model_entry.pack(side=tk.LEFT, padx=5)
        self.gemini_model_entry.insert(0, "gemini-2.5-pro")

        # vLLM Configuration
        self.vllm_frame = ttk.Frame(ai_frame)
        self.vllm_frame.grid(row=4, column=0, columnspan=2, sticky=tk.W+tk.E, pady=5)

        ttk.Label(self.vllm_frame, text="vLLM Base URL:").pack(side=tk.LEFT)
        self.vllm_url_entry = ttk.Entry(self.vllm_frame, width=50)
        self.vllm_url_entry.pack(side=tk.LEFT, padx=5)
        self.vllm_url_entry.insert(0, "http://localhost:8000/v1")

        ttk.Label(self.vllm_frame, text="Model:").pack(side=tk.LEFT, padx=(10, 0))
        self.vllm_model_entry = ttk.Entry(self.vllm_frame, width=30)
        self.vllm_model_entry.pack(side=tk.LEFT, padx=5)
        self.vllm_model_entry.insert(0, "Qwen/Qwen2-VL-72B-Instruct")

        # Initially hide all provider frames
        self.on_provider_changed(None)

        # Buttons
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(
            button_frame,
            text="Save Configuration",
            command=self.save_config
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="Load Configuration",
            command=self.load_config
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="Test Connection",
            command=self.test_connection
        ).pack(side=tk.LEFT, padx=5)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def on_provider_changed(self, event):
        """Show/hide provider-specific config based on selection"""
        provider = self.ai_provider_var.get()

        # Hide all frames
        self.claude_frame.grid_remove()
        self.openai_frame.grid_remove()
        self.gemini_frame.grid_remove()
        self.vllm_frame.grid_remove()

        # Show selected frame
        if provider == "claude":
            self.claude_frame.grid()
        elif provider == "openai":
            self.openai_frame.grid()
        elif provider == "gemini":
            self.gemini_frame.grid()
        elif provider == "vllm":
            self.vllm_frame.grid()

    def setup_about_tab(self):
        """Setup about tab"""

        about_text = """
        Trial Orders Automation
        Version 2.0 - Multi-Provider AI Support

        Automated processing of court trial orders with AI-powered date extraction.

        Features:
        • Automatic email monitoring via IMAP
        • PDF download and processing
        • Multi-provider AI support (Claude, OpenAI, Gemini, vLLM)
        • SharePoint integration via REST API
        • Case-specific filing
        • No Azure AD required

        Supported AI Providers:
        • Claude 3.5 Sonnet (Anthropic) - Best for document understanding
        • GPT-4o (OpenAI) - Strong vision capabilities
        • Gemini 1.5 Pro (Google) - Large context window
        • vLLM (Self-hosted) - Cost-effective, private, customizable
          - Supports: Qwen2-VL, LLaVA, CogVLM, and more

        vLLM Advantages:
        • Free inference after initial setup
        • Complete data privacy (runs on your hardware)
        • No API rate limits
        • Customizable models

        Technology Stack:
        • Python 3.9+
        • IMAP for email access
        • SharePoint REST API (Office365-REST-Python-Client)
        • Anthropic/OpenAI/Google APIs or vLLM

        For vLLM setup, see: https://docs.vllm.ai

        Internal use only
        For support, contact your IT administrator.
        """

        text_widget = tk.Text(
            self.tab_about,
            wrap=tk.WORD,
            font=("Segoe UI", 9),
            padx=20,
            pady=20
        )
        text_widget.insert("1.0", about_text)
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(fill=tk.BOTH, expand=True)

    def log(self, message, level="INFO"):
        """Add message to log queue"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put((timestamp, level, message))

    def check_log_queue(self):
        """Check log queue and update UI"""
        try:
            while True:
                timestamp, level, message = self.log_queue.get_nowait()

                # Insert into log text
                log_line = f"[{timestamp}] {message}\n"
                self.log_text.insert(tk.END, log_line, level)
                self.log_text.see(tk.END)

        except queue.Empty:
            pass

        # Schedule next check
        self.root.after(100, self.check_log_queue)

    def clear_log(self):
        """Clear the log text"""
        self.log_text.delete(1.0, tk.END)
        self.log("Log cleared", "INFO")

    def save_config(self):
        """Save configuration to file"""
        config = {
            "ai_provider": self.ai_provider_var.get(),
            "anthropic_key": self.anthropic_key_entry.get(),
            "openai_key": self.openai_key_entry.get(),
            "openai_model": self.openai_model_entry.get(),
            "gemini_key": self.gemini_key_entry.get(),
            "gemini_model": self.gemini_model_entry.get(),
            "vllm_url": self.vllm_url_entry.get(),
            "vllm_model": self.vllm_model_entry.get()
        }

        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)

            messagebox.showinfo("Success", "Configuration saved successfully!")
            self.log("Configuration saved", "SUCCESS")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {e}")
            self.log(f"Error saving configuration: {e}", "ERROR")

    def load_config(self):
        """Load configuration from file"""
        if not self.config_file.exists():
            self.log("No saved configuration found", "WARNING")
            return

        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)

            # AI config
            self.ai_provider_var.set(config.get("ai_provider", "claude"))
            self.on_provider_changed(None)

            self.anthropic_key_entry.delete(0, tk.END)
            self.anthropic_key_entry.insert(0, config.get("anthropic_key", ""))

            self.openai_key_entry.delete(0, tk.END)
            self.openai_key_entry.insert(0, config.get("openai_key", ""))

            self.openai_model_entry.delete(0, tk.END)
            self.openai_model_entry.insert(0, config.get("openai_model", "gpt-4o"))

            self.gemini_key_entry.delete(0, tk.END)
            self.gemini_key_entry.insert(0, config.get("gemini_key", ""))

            self.gemini_model_entry.delete(0, tk.END)
            self.gemini_model_entry.insert(0, config.get("gemini_model", "gemini-2.5-pro"))

            self.vllm_url_entry.delete(0, tk.END)
            self.vllm_url_entry.insert(0, config.get("vllm_url", "http://localhost:8000/v1"))

            self.vllm_model_entry.delete(0, tk.END)
            self.vllm_model_entry.insert(0, config.get("vllm_model", "Qwen/Qwen2-VL-72B-Instruct"))

            self.log("Configuration loaded", "SUCCESS")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load configuration: {e}")
            self.log(f"Error loading configuration: {e}", "ERROR")

    def test_connection(self):
        """Test the connection to Outlook and OneDrive folders"""
        self.log("Testing connection...", "INFO")

        # Set environment variables
        self.set_env_vars()

        try:
            from automation import EmailClient, SharePointClient, Config

            config = Config()

            # Test Outlook
            self.log("Testing Outlook connection...", "INFO")
            email_client = EmailClient(config)
            email_client.connect()
            email_client.disconnect()
            self.log("Outlook connection successful", "SUCCESS")

            # Test OneDrive folders
            self.log("Testing OneDrive folders...", "INFO")
            sp_client = SharePointClient(config)
            self.log("OneDrive folders accessible", "SUCCESS")

            messagebox.showinfo("Success", "All connections successful!\n\nOutlook: ✅\nOneDrive: ✅")

        except Exception as e:
            messagebox.showerror("Connection Failed", str(e))
            self.log(f"Connection test failed: {e}", "ERROR")

    def set_env_vars(self):
        """Set environment variables from GUI inputs"""
        os.environ["AI_PROVIDER"] = self.ai_provider_var.get()
        os.environ["ANTHROPIC_API_KEY"] = self.anthropic_key_entry.get()
        os.environ["OPENAI_API_KEY"] = self.openai_key_entry.get()
        os.environ["OPENAI_MODEL"] = self.openai_model_entry.get()
        os.environ["GEMINI_API_KEY"] = self.gemini_key_entry.get()
        os.environ["GEMINI_MODEL"] = self.gemini_model_entry.get()
        os.environ["VLLM_BASE_URL"] = self.vllm_url_entry.get()
        os.environ["VLLM_MODEL"] = self.vllm_model_entry.get()

    def start_processing(self):
        """Start the automation process"""

        # Validate basic configuration
        provider = self.ai_provider_var.get()

        # Validate AI provider
        if provider == "claude" and not self.anthropic_key_entry.get():
            messagebox.showwarning("API Key Required", "Please enter Anthropic API key")
            return
        elif provider == "openai" and not self.openai_key_entry.get():
            messagebox.showwarning("API Key Required", "Please enter OpenAI API key")
            return
        elif provider == "gemini" and not self.gemini_key_entry.get():
            messagebox.showwarning("API Key Required", "Please enter Gemini API key")
            return
        elif provider == "vllm" and not self.vllm_url_entry.get():
            messagebox.showwarning("Configuration Required", "Please enter vLLM base URL")
            return

        # Set environment variables
        self.set_env_vars()

        # Update UI
        self.is_running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress.start()
        self.status_label.config(text=f"Processing emails using {provider.upper()}...")

        # Start processing in background thread
        self.process_thread = threading.Thread(target=self.run_automation, daemon=True)
        self.process_thread.start()

        self.log("=" * 60, "INFO")
        self.log(f"Started automation process (AI: {provider.upper()})", "INFO")
        self.log("=" * 60, "INFO")

    def stop_processing(self):
        """Stop the automation process"""
        self.is_running = False
        self.stop_button.config(state=tk.DISABLED)
        self.log("Stopping... (will complete current task)", "WARNING")

    def run_automation(self):
        """Run the automation in background thread"""

        try:
            # Redirect print statements to log
            class LogCapture:
                def write(self, message):
                    if message.strip():
                        if "✅" in message or "SUCCESS" in message.upper():
                            level = "SUCCESS"
                        elif "⚠️" in message or "WARNING" in message.upper():
                            level = "WARNING"
                        elif "❌" in message or "ERROR" in message.upper():
                            level = "ERROR"
                        else:
                            level = "INFO"

                        self.gui.log(message.strip(), level)

                def flush(self):
                    pass

            log_capture = LogCapture()
            log_capture.gui = self
            sys.stdout = log_capture

            # Run automation
            from automation import TrialOrdersAutomation
            automation = TrialOrdersAutomation()
            automation.run()

            # Update stats
            self.root.after(0, lambda: self.last_run_label.config(
                text=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ))

            self.log("=" * 60, "INFO")
            self.log("Automation completed successfully", "SUCCESS")
            self.log("=" * 60, "INFO")

        except Exception as e:
            self.log(f"Fatal error: {e}", "ERROR")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Automation failed: {e}"))

        finally:
            # Restore stdout
            sys.stdout = sys.__stdout__

            # Update UI
            self.root.after(0, self.finish_processing)

    def finish_processing(self):
        """Finish processing and update UI"""
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.progress.stop()
        self.status_label.config(text="Ready to process emails")


def main():
    """Main entry point"""
    root = tk.Tk()
    app = AutomationGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
