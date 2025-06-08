import tkinter as tk
from tkinter import ttk
import threading
import time
from pathlib import Path
import os

class SplashScreen:
    def __init__(self, title="Obsidian Vault Health Check", version="v1.0"):
        self.title = title
        self.version = version
        self.root = None
        self.status_label = None
        self.progress_var = None
        self.progress_bar = None
        self.is_closed = False
        
    def create_splash(self):
        """Create and display the splash screen"""
        self.root = tk.Tk()
        self.root.title(self.title)
        
        # Set window properties
        window_width = 400
        window_height = 300
        
        # Center the window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(False, False)
        self.root.configure(bg='#2b2b2b')  # Dark background
        
        # Remove window decorations for splash effect
        self.root.overrideredirect(True)
        
        # Create main frame
        main_frame = tk.Frame(self.root, bg='#2b2b2b', relief='raised', bd=2)
        main_frame.pack(fill='both', expand=True, padx=2, pady=2)
        
        # Create logo area (circular placeholder)
        logo_frame = tk.Frame(main_frame, bg='#2b2b2b')
        logo_frame.pack(pady=(40, 20))
        
        # Create a circular logo using Canvas
        canvas = tk.Canvas(logo_frame, width=100, height=100, bg='#2b2b2b', highlightthickness=0)
        canvas.pack()
        
        # Draw circular logo
        canvas.create_oval(10, 10, 90, 90, fill='#7c3aed', outline='#a855f7', width=3)
        canvas.create_text(50, 50, text='OV', font=('Arial', 24, 'bold'), fill='white')
        
        # Title and version
        title_label = tk.Label(main_frame, text=self.title, 
                              font=('Arial', 16, 'bold'), 
                              fg='white', bg='#2b2b2b')
        title_label.pack(pady=(10, 5))
        
        version_label = tk.Label(main_frame, text=self.version, 
                               font=('Arial', 10), 
                               fg='#9ca3af', bg='#2b2b2b')
        version_label.pack()
        
        # Progress bar
        progress_frame = tk.Frame(main_frame, bg='#2b2b2b')
        progress_frame.pack(pady=(30, 10), padx=40, fill='x')
        
        self.progress_var = tk.DoubleVar()
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Custom.Horizontal.TProgressbar", 
                       background='#7c3aed',
                       troughcolor='#374151',
                       borderwidth=0,
                       lightcolor='#7c3aed',
                       darkcolor='#7c3aed')
        
        self.progress_bar = ttk.Progressbar(progress_frame, 
                                          variable=self.progress_var,
                                          maximum=100,
                                          style="Custom.Horizontal.TProgressbar",
                                          length=300)
        self.progress_bar.pack(fill='x')
        
        # Status text (bottom left)
        status_frame = tk.Frame(main_frame, bg='#2b2b2b')
        status_frame.pack(fill='x', side='bottom', pady=(0, 20), padx=20)
        
        self.status_label = tk.Label(status_frame, text="Initializing...", 
                                   font=('Arial', 9), 
                                   fg='#9ca3af', bg='#2b2b2b',
                                   anchor='w')
        self.status_label.pack(fill='x')
        
        # Make sure the splash is on top
        self.root.lift()
        self.root.attributes('-topmost', True)
        
        return self.root
    
    def update_status(self, message, progress=None):
        """Update the status message and optionally the progress"""
        if self.is_closed or not self.root:
            return
            
        try:
            if self.status_label:
                self.status_label.config(text=message)
            
            if progress is not None and self.progress_var:
                self.progress_var.set(progress)
            
            self.root.update_idletasks()
            self.root.update()
        except tk.TclError:
            # Window was closed
            self.is_closed = True
    
    def close(self):
        """Close the splash screen"""
        if self.root and not self.is_closed:
            self.is_closed = True
            try:
                self.root.destroy()
            except tk.TclError:
                pass  # Already destroyed
    
    def show_and_run(self, target_function, *args, **kwargs):
        """Show splash screen and run target function in background"""
        # Create splash screen
        splash_root = self.create_splash()
        
        # Result container
        result_container = {'result': None, 'exception': None, 'completed': False}
        
        def run_target():
            """Run the target function in a separate thread"""
            try:
                result_container['result'] = target_function(*args, **kwargs)
            except Exception as e:
                result_container['exception'] = e
            finally:
                result_container['completed'] = True
        
        # Start the target function in a background thread
        thread = threading.Thread(target=run_target, daemon=True)
        thread.start()
        
        # Keep splash screen alive until target completes
        while not result_container['completed'] and not self.is_closed:
            try:
                splash_root.update()
                time.sleep(0.1)  # Small delay to prevent high CPU usage
            except tk.TclError:
                # Window was closed by user
                self.is_closed = True
                break
        
        # Close splash screen
        self.close()
        
        # Handle any exceptions that occurred
        if result_container['exception']:
            raise result_container['exception']
        
        return result_container['result']


class ProgressReporter:
    """Helper class to report progress to splash screen"""
    def __init__(self, splash_screen, total_steps=100):
        self.splash = splash_screen
        self.total_steps = total_steps
        self.current_step = 0
        
    def update(self, message, step_increment=1):
        """Update progress with message and increment step"""
        self.current_step += step_increment
        progress = min((self.current_step / self.total_steps) * 100, 100)
        if self.splash:
            self.splash.update_status(message, progress)
        
    def set_progress(self, message, progress_percent):
        """Set specific progress percentage"""
        if self.splash:
            self.splash.update_status(message, progress_percent)
    
    def complete(self, message="Complete"):
        """Mark as complete"""
        if self.splash:
            self.splash.update_status(message, 100)