import cv2
import os
import time
import threading
from datetime import timedelta
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

class VideoToImageApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Video to Image Converter")
        self.root.geometry("800x600")
        self.root.minsize(600, 500)  # Smaller minimum size that still fits all elements
        self.root.resizable(True, True)
        
        # Bind resize event to adjust UI elements when window size changes
        self.root.bind("<Configure>", self.on_window_resize)
        
        # Initialize variables
        self.video_source = None
        self.output_folder = None
        self.interval = ttk.DoubleVar(value=1.0)
        self.frame_count = ttk.IntVar(value=10)
        self.output_format = ttk.StringVar(value="jpg")
        self.extraction_method = ttk.StringVar(value="interval")
        self.is_camera = False
        self.camera_idx = None
        self.preview_running = False
        self.preview_thread = None
        self.cap = None
        self.total_frames = 0
        self.progress_value = ttk.DoubleVar(value=0)
        self.status_text = ttk.StringVar(value="Ready")
        
        # Create main frames
        self.create_widgets()
        
        # Bind closing event
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def create_widgets(self):
        # Create a notebook with tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Main tab
        main_frame = ttk.Frame(notebook)
        notebook.add(main_frame, text="Main")
        
        # About tab
        about_frame = ttk.Frame(notebook)
        notebook.add(about_frame, text="About")
        
        # Setup main tab
        self.setup_main_tab(main_frame)
        
        # Setup about tab
        self.setup_about_tab(about_frame)
    
    def setup_main_tab(self, parent):
        # Create a main container frame that uses grid for better spacing
        main_container = ttk.Frame(parent)
        main_container.pack(fill="both", expand=True)
        
        # Configure row and column weights to make the layout responsive
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(3, weight=1)  # Preview frame gets the extra space
        
        # Create frames using grid layout
        source_frame = ttk.LabelFrame(main_container, text="Source Selection", padding=(5, 5))
        source_frame.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")
        
        output_frame = ttk.LabelFrame(main_container, text="Output Settings", padding=(5, 5))
        output_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        
        extraction_frame = ttk.LabelFrame(main_container, text="Extraction Method", padding=(5, 5))
        extraction_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")
        
        preview_frame = ttk.LabelFrame(main_container, text="Preview", padding=(5, 5))
        preview_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        
        control_frame = ttk.Frame(main_container)
        control_frame.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")
        
        # Configure source frame columns
        source_frame.columnconfigure(0, weight=1)
        source_frame.columnconfigure(1, weight=1)
        
        # Source Selection with more compact layout
        button_frame = ttk.Frame(source_frame)
        button_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        
        file_button = ttk.Button(
            button_frame, 
            text="Select Video File", 
            command=self.select_video_file,
            bootstyle=SUCCESS
        )
        file_button.grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        camera_button = ttk.Button(
            button_frame, 
            text="Select Camera", 
            command=self.select_camera,
            bootstyle=INFO
        )
        camera_button.grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        
        self.source_label = ttk.Label(source_frame, text="No source selected")
        self.source_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        
        # Configure output frame columns
        output_frame.columnconfigure(0, weight=2)
        output_frame.columnconfigure(1, weight=1)
        output_frame.columnconfigure(2, weight=1)
        
        # Output Settings with more compact layout
        output_button = ttk.Button(
            output_frame, 
            text="Select Output Folder", 
            command=self.select_output_directory,
            bootstyle=PRIMARY
        )
        output_button.grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        format_label = ttk.Label(output_frame, text="Format:")
        format_label.grid(row=0, column=1, padx=2, pady=2, sticky="e")
        
        format_combo = ttk.Combobox(
            output_frame, 
            textvariable=self.output_format,
            values=["jpg", "png"],
            width=4,
            state="readonly"
        )
        format_combo.grid(row=0, column=2, padx=2, pady=2, sticky="w")
        
        self.output_label = ttk.Label(output_frame, text="No output folder selected")
        self.output_label.grid(row=1, column=0, columnspan=3, sticky="w", padx=5, pady=2)
        
        # Extraction Method
        interval_radio = ttk.Radiobutton(
            extraction_frame,
            text="Extract by Time Interval",
            variable=self.extraction_method,
            value="interval",
            command=self.update_extraction_options
        )
        interval_radio.grid(row=0, column=0, padx=2, pady=2, sticky="w")
        
        count_radio = ttk.Radiobutton(
            extraction_frame,
            text="Extract by Frame Count",
            variable=self.extraction_method,
            value="count",
            command=self.update_extraction_options
        )
        count_radio.grid(row=1, column=0, padx=2, pady=2, sticky="w")
        
        # Configure extraction frame columns
        extraction_frame.columnconfigure(0, weight=1)
        extraction_frame.columnconfigure(1, weight=1)
        
        # Interval options - use grid instead of pack for better alignment
        self.interval_frame = ttk.Frame(extraction_frame)
        self.interval_frame.grid(row=0, column=1, padx=2, pady=2, sticky="w")
        self.interval_frame.columnconfigure(0, weight=0)
        self.interval_frame.columnconfigure(1, weight=0)
        
        interval_label = ttk.Label(self.interval_frame, text="Interval (s):")
        interval_label.grid(row=0, column=0, padx=2, sticky="e")
        
        interval_entry = ttk.Spinbox(
            self.interval_frame,
            from_=0.1,
            to=60,
            increment=0.1,
            textvariable=self.interval,
            width=4
        )
        interval_entry.grid(row=0, column=1, padx=2, sticky="w")
        
        # Count options - use grid instead of pack for better alignment
        self.count_frame = ttk.Frame(extraction_frame)
        self.count_frame.grid(row=1, column=1, padx=2, pady=2, sticky="w")
        self.count_frame.columnconfigure(0, weight=0)
        self.count_frame.columnconfigure(1, weight=0)
        
        count_label = ttk.Label(self.count_frame, text="Frames:")
        count_label.grid(row=0, column=0, padx=2, sticky="e")
        
        count_entry = ttk.Spinbox(
            self.count_frame,
            from_=1,
            to=1000,
            increment=1,
            textvariable=self.frame_count,
            width=4
        )
        count_entry.grid(row=0, column=1, padx=2, sticky="w")
        
        # Set initial state
        self.update_extraction_options()
        
        # Configure preview frame for responsive sizing
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # Preview canvas with responsive sizing
        self.preview_canvas = ttk.Canvas(preview_frame, bg="black")
        self.preview_canvas.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        # Add preview button directly to the frame
        preview_controls = ttk.Frame(preview_frame)
        preview_controls.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        
        self.preview_button = ttk.Button(
            preview_controls,
            text="Start Preview",
            command=self.toggle_preview,
            bootstyle=OUTLINE
        )
        self.preview_button.pack(side="left", padx=5)
        
        # Control frame with grid layout for better space distribution
        control_frame.columnconfigure(0, weight=2)  # Progress bar gets more space
        control_frame.columnconfigure(1, weight=1)  # Status
        control_frame.columnconfigure(2, weight=0)  # Button
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            control_frame, 
            variable=self.progress_value,
            bootstyle=SUCCESS
        )
        self.progress_bar.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        # Status label
        self.status_label = ttk.Label(control_frame, textvariable=self.status_text)
        self.status_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # Extract button
        self.extract_button = ttk.Button(
            control_frame,
            text="Extract Frames",
            command=self.start_extraction,
            bootstyle=SUCCESS
        )
        self.extract_button.grid(row=0, column=2, padx=5, pady=5, sticky="e")
    
    def setup_about_tab(self, parent):
        # Use grid layout instead of pack for better control
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        
        about_label = ttk.Label(
            parent,
            text="Video to Image Converter",
            font=("Helvetica", 14, "bold")
        )
        about_label.grid(row=0, column=0, pady=(15, 5))
        
        description = (
            "Extract image frames from videos or camera feeds.\n\n"
            "• Video files and camera inputs\n"
            "• Extract by time interval or frame count\n"
            "• JPG or PNG output formats\n"
            "• Real-time preview\n\n"
            "Created with ttkbootstrap and OpenCV"
        )
        
        desc_label = ttk.Label(
            parent,
            text=description,
            justify="center",
            wraplength=400
        )
        desc_label.grid(row=1, column=0, padx=20, pady=10, sticky="n")
    
    def update_extraction_options(self):
        method = self.extraction_method.get()
        if method == "interval":
            self.interval_frame.configure(style='TFrame')
            self.count_frame.configure(style='muted.TFrame')
        else:
            self.interval_frame.configure(style='muted.TFrame')
            self.count_frame.configure(style='TFrame')
    
    def on_window_resize(self, event):
        """Handle window resize events to adjust UI elements."""
        # Only respond to root window resizing, not internal widgets
        if event.widget == self.root:
            # Update preview area dimensions
            self.adjust_preview_size()
    
    def adjust_preview_size(self):
        """Adjust the preview size based on the current window dimensions."""
        # Get window dimensions
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        
        # Calculate appropriate preview height (about 40% of window height)
        preview_height = int(window_height * 0.4)
        
        # Ensure there's still room for other elements
        if preview_height < 150:
            preview_height = 150
        elif preview_height > 400:
            preview_height = 400
    
    def select_video_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Video File",
            filetypes=[
                ("Video files", "*.mp4 *.avi *.mov *.mkv *.wmv *.flv"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.video_source = file_path
            self.is_camera = False
            self.source_label.configure(text=f"Selected: {os.path.basename(file_path)}")
            self.status_text.set(f"Video file selected: {os.path.basename(file_path)}")
            
            # Get video info
            self.get_video_info()
    
    def select_camera(self):
        # Find available cameras
        available_cameras = self.find_available_cameras()
        
        if not available_cameras:
            messagebox.showinfo("No Cameras", "No camera devices found.")
            return
        
        # Create camera selection dialog
        camera_dialog = ttk.Toplevel(self.root)
        camera_dialog.title("Select Camera")
        camera_dialog.geometry("300x200")
        camera_dialog.resizable(False, False)
        camera_dialog.transient(self.root)
        camera_dialog.grab_set()
        
        ttk.Label(
            camera_dialog, 
            text="Select a camera:",
            font=("Helvetica", 12)
        ).pack(pady=10)
        
        camera_var = ttk.IntVar()
        
        for idx, name in available_cameras.items():
            ttk.Radiobutton(
                camera_dialog,
                text=name,
                variable=camera_var,
                value=idx
            ).pack(anchor="w", padx=20, pady=5)
        
        def on_select():
            self.camera_idx = camera_var.get()
            self.video_source = self.camera_idx
            self.is_camera = True
            self.source_label.configure(text=f"Selected: {available_cameras[self.camera_idx]}")
            self.status_text.set(f"Camera selected: {available_cameras[self.camera_idx]}")
            camera_dialog.destroy()
        
        ttk.Button(
            camera_dialog,
            text="Select",
            command=on_select,
            bootstyle=SUCCESS
        ).pack(pady=10)
        
        # Center the dialog
        camera_dialog.update_idletasks()
        width = camera_dialog.winfo_width()
        height = camera_dialog.winfo_height()
        x = (camera_dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (camera_dialog.winfo_screenheight() // 2) - (height // 2)
        camera_dialog.geometry(f'{width}x{height}+{x}+{y}')
        
        self.root.wait_window(camera_dialog)
    
    def find_available_cameras(self):
        available_cameras = {}
        max_test_port = 10
        
        self.status_text.set("Searching for available camera devices...")
        
        for i in range(max_test_port):
            cap = cv2.VideoCapture(i)
            if cap.isOpened():
                # Try to get the camera name
                ret, frame = cap.read()
                if ret:
                    camera_name = f"Camera {i}"
                    available_cameras[i] = camera_name
                cap.release()
        
        self.status_text.set(f"Found {len(available_cameras)} camera devices")
        return available_cameras
    
    def select_output_directory(self):
        directory = filedialog.askdirectory(
            title="Select Output Directory",
            mustexist=True
        )
        
        if directory:
            self.output_folder = directory
            self.output_label.configure(text=f"Selected: {directory}")
            self.status_text.set(f"Output folder selected: {directory}")
    
    def get_video_info(self):
        if not self.is_camera and self.video_source:
            cap = cv2.VideoCapture(self.video_source)
            if cap.isOpened():
                # Get video properties
                self.total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
                fps = cap.get(cv2.CAP_PROP_FPS)
                width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                
                duration = self.total_frames / fps if fps > 0 else 0
                
                info_text = f"Video Info: {width}x{height}, {fps:.2f} FPS, Duration: {timedelta(seconds=int(duration))}, Frames: {self.total_frames}"
                self.status_text.set(info_text)
                
                # Release the capture
                cap.release()
    
    def toggle_preview(self):
        if self.preview_running:
            self.preview_running = False
            self.preview_button.configure(text="Start Preview")
            if self.cap is not None:
                self.cap.release()
                self.cap = None
        else:
            if self.video_source is not None:
                self.preview_running = True
                self.preview_button.configure(text="Stop Preview")
                
                # Clear any previous content
                self.preview_canvas.delete("all")
                
                # Show loading message
                center_x = self.preview_canvas.winfo_width() // 2 or 320
                center_y = self.preview_canvas.winfo_height() // 2 or 180
                
                self.preview_canvas.create_text(
                    center_x, center_y - 15,
                    text="Loading preview...",
                    fill="white",
                    font=("Helvetica", 12)
                )
                
                # Add a hint about saving frames from camera
                if self.is_camera:
                    self.preview_canvas.create_text(
                        center_x, center_y + 15,
                        text="For camera capture: press 'c' to capture frame, 'q' to stop",
                        fill="#aaaaaa",
                        font=("Helvetica", 9)
                    )
                
                self.root.update_idletasks()
                
                # Start preview in a separate thread
                self.preview_thread = threading.Thread(target=self.run_preview)
                self.preview_thread.daemon = True
                self.preview_thread.start()
            else:
                messagebox.showinfo("Error", "Please select a video source first.")
    
    def run_preview(self):
        self.cap = cv2.VideoCapture(self.video_source)
        
        if not self.cap.isOpened():
            messagebox.showerror("Error", "Could not open video source.")
            self.preview_running = False
            self.preview_button.configure(text="Start Preview")
            return
        
        # Get the first frame to determine aspect ratio
        ret, frame = self.cap.read()
        if ret:
            # Reset to beginning for video files
            if not self.is_camera:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                
            # Get video dimensions for proper scaling
            height, width = frame.shape[:2]
            self.video_aspect_ratio = width / height
            
            # Update UI for this aspect ratio
            self.root.after(10, self.apply_aspect_ratio)
        
        while self.preview_running:
            ret, frame = self.cap.read()
            
            if not ret:
                # If it's a video file and we reached the end, restart
                if not self.is_camera:
                    self.cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                    continue
                else:
                    break
            
            # Resize frame to fit canvas
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            
            if canvas_width > 1 and canvas_height > 1:
                frame = self.resize_frame(frame, canvas_width, canvas_height)
            else:
                # Use default dimensions if canvas is not yet properly sized
                frame = self.resize_frame(frame, 640, 360)
                
            # Convert frame to PhotoImage
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame_rgb)
            photo = ImageTk.PhotoImage(image=img)
            
            # Clear previous content and update canvas
            self.preview_canvas.delete("all")
            
            # Determine center coordinates (handle both zero and non-zero dimensions)
            center_x = max(canvas_width // 2, 1)
            center_y = max(canvas_height // 2, 1)
            
            # Update canvas
            self.preview_canvas.create_image(
                center_x, 
                center_y, 
                image=photo, 
                anchor="center"
            )
            self.preview_canvas.image = photo  # Keep a reference
            
            # Process any pending events to keep UI responsive
            self.root.update_idletasks()
            
            # Need a small delay to refresh UI
            time.sleep(0.03)
        
        if self.cap is not None:
            self.cap.release()
            self.cap = None

    def apply_aspect_ratio(self):
        """Apply the video's aspect ratio to the preview canvas if possible"""
        if hasattr(self, 'video_aspect_ratio'):
            # Get the available width
            preview_width = self.preview_canvas.winfo_width()
            if preview_width > 10:  # Only proceed if the canvas has a reasonable width
                # Calculate the appropriate height based on aspect ratio
                preview_height = int(preview_width / self.video_aspect_ratio)
                
                # Make sure it's not too tall or too short
                min_height = 150
                max_height = 350
                if preview_height < min_height:
                    preview_height = min_height
                elif preview_height > max_height:
                    preview_height = max_height
                
                # Get the preview frame
                preview_frame = self.preview_canvas.master
                
                # Configure the grid row for the preview frame
                parent_grid = preview_frame.master
                parent_grid.rowconfigure(3, minsize=preview_height)
    
    def resize_frame(self, frame, target_width, target_height):
        height, width = frame.shape[:2]
        
        # Calculate aspect ratios
        aspect_ratio = width / height
        target_ratio = target_width / target_height
        
        # Determine new dimensions
        if aspect_ratio > target_ratio:
            # Width-constrained
            new_width = target_width
            new_height = int(target_width / aspect_ratio)
        else:
            # Height-constrained
            new_height = target_height
            new_width = int(target_height * aspect_ratio)
        
        # Resize frame
        resized = cv2.resize(frame, (new_width, new_height), interpolation=cv2.INTER_AREA)
        
        return resized
    
    def start_extraction(self):
        if not self.video_source:
            messagebox.showerror("Error", "Please select a video source.")
            return
        
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder.")
            return
        
        # Stop preview if running
        if self.preview_running:
            self.toggle_preview()
        
        # Start extraction in a separate thread
        extraction_thread = threading.Thread(target=self.run_extraction)
        extraction_thread.daemon = True
        extraction_thread.start()
    
    def run_extraction(self):
        self.extract_button.configure(state="disabled")
        method = self.extraction_method.get()
        
        try:
            if method == "interval":
                self.extract_frames_by_interval()
            else:
                self.extract_frames_by_count()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during extraction: {str(e)}")
        finally:
            self.extract_button.configure(state="normal")
    
    def extract_frames_by_interval(self):
        interval_seconds = self.interval.get()
        output_format = self.output_format.get()
        
        # Create output folder if it doesn't exist
        os.makedirs(self.output_folder, exist_ok=True)
        
        # Open the video file or camera
        video = cv2.VideoCapture(self.video_source)
        
        if not video.isOpened():
            messagebox.showerror("Error", f"Could not open video source {self.video_source}")
            return
        
        # Get video properties
        fps = video.get(cv2.CAP_PROP_FPS)
        
        # For camera input, fps may be low or unreliable, so we handle that differently
        if self.is_camera or fps < 0.1:
            fps = 30  # Assume 30 fps for cameras
            self.status_text.set(f"Using estimated frame rate: {fps} fps")
            
            frame_number = 0
            start_time = time.time()
            
            while True:
                # Read the frame
                ret, frame = video.read()
                
                # Break the loop if we can't read any more frames
                if not ret:
                    break
                
                current_time = time.time()
                elapsed_time = current_time - start_time
                
                # Check if it's time to save a frame
                if elapsed_time >= frame_number * interval_seconds:
                    # Save the frame as an image
                    timestamp_str = str(timedelta(seconds=int(elapsed_time))).replace(':', '-')
                    output_file = os.path.join(self.output_folder, f"frame_{frame_number:04d}_{timestamp_str}.{output_format}")
                    cv2.imwrite(output_file, frame)
                    
                    self.status_text.set(f"Saved frame #{frame_number}")
                    frame_number += 1
                
                # Update progress (arbitrary for camera mode)
                self.progress_value.set(min(frame_number * 10, 100))
                
                # Process events to keep UI responsive
                self.root.update_idletasks()
                
                # Check if cancel requested
                if not self.extract_button.instate(['disabled']):
                    break
        else:
            # For video files, use the frame-based approach
            frame_count = int(video.get(cv2.CAP_PROP_FRAME_COUNT))
            duration = frame_count / fps
            
            info_text = f"Extracting frames ({interval_seconds}s interval)"
            self.status_text.set(info_text)
            
            # Calculate frame interval
            frame_interval = int(fps * interval_seconds)
            
            current_frame = 0
            frame_number = 0
            
            while True:
                # Set video position to the desired frame
                video.set(cv2.CAP_PROP_POS_FRAMES, current_frame)
                
                # Read the frame
                ret, frame = video.read()
                
                # Break the loop if we can't read any more frames
                if not ret:
                    break
                
                # Save the frame as an image
                timestamp = current_frame / fps
                timestamp_str = str(timedelta(seconds=int(timestamp))).replace(':', '-')
                output_file = os.path.join(self.output_folder, f"frame_{frame_number:04d}_{timestamp_str}.{output_format}")
                cv2.imwrite(output_file, frame)
                
                self.status_text.set(f"Saved frame #{frame_number}")
                
                # Move to the next frame
                current_frame += frame_interval
                frame_number += 1
                
                # Update progress
                progress = min(current_frame / frame_count * 100, 100)
                self.progress_value.set(progress)
                
                # Process events to keep UI responsive
                self.root.update_idletasks()
                
                # Break if we've reached the end of the video
                if current_frame >= frame_count:
                    break
                
                # Check if cancel requested
                if not self.extract_button.instate(['disabled']):
                    break
        
        # Clean up
        video.release()
        
        # Final status update
        self.progress_value.set(100)
        self.status_text.set(f"Extracted {frame_number} frames")
        
        # Show completion message
    
    def extract_frames_by_count(self):
        total_frames = self.frame_count.get()
        output_format = self.output_format.get()
        
        # Create output folder if it doesn't exist
        os.makedirs(self.output_folder, exist_ok=True)
        
        # Open the video file or camera
        video = cv2.VideoCapture(self.video_source)
        
        if not video.isOpened():
            messagebox.showerror("Error", f"Could not open video source {self.video_source}")
            return
        
        # Check if it's a camera (integer index) or a file
        if self.is_camera:
            messagebox.showinfo("Camera Mode", 
                               "Camera mode: Press 'c' in the preview window to capture a frame, 'q' to stop.")
            
            frames_captured = 0
            
            # Create a preview window
            cv2.namedWindow('Camera Capture', cv2.WINDOW_NORMAL)
            
            while frames_captured < total_frames:
                # Read the frame
                ret, frame = video.read()
                
                # Break the loop if we can't read the frame
                if not ret:
                    break
                
                # Display frame with instructions
                display_frame = frame.copy()
                cv2.putText(display_frame, f"Press 'c' to capture ({frames_captured}/{total_frames})", 
                           (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
                
                cv2.imshow('Camera Capture', display_frame)
                
                # Wait for a key press
                key = cv2.waitKey(1) & 0xFF
                
                # If 'c' is pressed, capture the frame
                if key == ord('c'):
                    output_file = os.path.join(self.output_folder, f"frame_{frames_captured:04d}.{output_format}")
                    cv2.imwrite(output_file, frame)
                    
                    frames_captured += 1
                    self.status_text.set(f"Captured {frames_captured}/{total_frames} frames")
                    self.progress_value.set(frames_captured / total_frames * 100)
                
                # If 'q' is pressed, quit
                elif key == ord('q'):
                    break
                
                # Check if cancel requested
                if not self.extract_button.instate(['disabled']):
                    break
            
            cv2.destroyWindow('Camera Capture')
        else:
            # For video files
            frame_count = int(video.get(cv2.CAP_PROP_FRAME_COUNT))
            fps = video.get(cv2.CAP_PROP_FPS)
            duration = frame_count / fps
            
            info_text = (f"Extracting {total_frames} evenly spaced frames from "
                        f"{timedelta(seconds=int(duration))} video with {frame_count} frames")
            self.status_text.set(info_text)
            
            # Calculate frame interval
            if total_frames > frame_count:
                total_frames = frame_count
                self.status_text.set(f"Warning: Video has fewer frames than requested. Extracting all {total_frames} frames.")
            
            frame_interval = frame_count / total_frames
            
            for i in range(total_frames):
                # Calculate the frame position
                frame_position = int(i * frame_interval)
                
                # Set video position to the desired frame
                video.set(cv2.CAP_PROP_POS_FRAMES, frame_position)
                
                # Read the frame
                ret, frame = video.read()
                
                # Break the loop if we can't read any more frames
                if not ret:
                    break
                
                # Save the frame as an image
                timestamp = frame_position / fps
                timestamp_str = str(timedelta(seconds=int(timestamp))).replace(':', '-')
                output_file = os.path.join(self.output_folder, f"frame_{i:04d}_{timestamp_str}.{output_format}")
                cv2.imwrite(output_file, frame)
                
                self.status_text.set(f"Saved: {os.path.basename(output_file)}")
                
                # Update progress
                progress = (i + 1) / total_frames * 100
                self.progress_value.set(progress)
                
                # Process events to keep UI responsive
                self.root.update_idletasks()
                
                # Check if cancel requested
                if not self.extract_button.instate(['disabled']):
                    break
        
        # Clean up
        video.release()
        cv2.destroyAllWindows()
        
        # Final status update
        self.progress_value.set(100)
        self.status_text.set(f"Extracted {total_frames} frames to {self.output_folder}")
        
        # Show completion message
        messagebox.showinfo("Extraction Complete", f"Successfully extracted {total_frames} frames to {self.output_folder}")
    
    def on_close(self):
        # Stop preview if running
        if self.preview_running:
            self.preview_running = False
            if self.cap is not None:
                self.cap.release()
        
        # Close all cv2 windows
        cv2.destroyAllWindows()
        
        # Close the application
        self.root.destroy()

if __name__ == "__main__":
    # Setup exception handler for better error reporting
    def show_error(exc_type, exc_value, exc_tb):
        import traceback
        traceback_details = '\n'.join(traceback.format_exception(exc_type, exc_value, exc_tb))
        messagebox.showerror(
            "Error",
            f"An unexpected error occurred:\n\n{exc_value}"
        )
        
        # Log the error
        with open("error_log.txt", "a") as f:
            f.write(f"--- Error: {time.strftime('%Y-%m-%d %H:%M:%S')} ---\n")
            f.write(traceback_details)
            f.write("\n\n")
    
    # Create the application
    try:
        # Create root window with theme
        root = ttk.Window(themename="darkly")
        
        # Center the window on screen
        def center_window(win, width=800, height=600):
            # Get screen dimensions
            screen_width = win.winfo_screenwidth()
            screen_height = win.winfo_screenheight()
            
            # Calculate position coordinates
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            
            # Set the window position
            win.geometry(f"{width}x{height}+{x}+{y}")
        
        center_window(root)
        
        # Try to set application icon
        try:
            root.iconbitmap("video_to_image.ico")
        except:
            pass  # Icon file not found, use default
            
        app = VideoToImageApp(root)
        
        # Set exception handler
        import sys
        sys.excepthook = show_error
        
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Startup Error", f"Failed to start the application: {str(e)}")
        raise