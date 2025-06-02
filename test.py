# Import required libraries
import tkinter as tk  # For creating the GUI interface
from tkinter import messagebox, ttk  # For popup messages and improved widgets
from pymodbus.client.sync import ModbusTcpClient  # For communicating with ADAM device
import logging  # For recording program activity
import threading  # For running background tasks
import time  # For adding delays

# Set up basic logging configuration
logging.basicConfig()  # Initialize logging system
logger = logging.getLogger()  # Create logger object
logger.setLevel(logging.INFO)  # Set logging level to INFO (not DEBUG or WARNING)

# Device connection settings
ADAM_IP = "192.168.10.15"  # IP address of the ADAM-6224 module
PORT = 502  # Modbus TCP port number (standard for Modbus)
VERIFICATION_INTERVAL = 1.0  # How often to check values (in seconds)

# Map each channel to its Modbus register address
AO_CHANNEL_REGISTERS = {
    0: 0,  # Channel 0 uses register 0
    1: 1,  # Channel 1 uses register 1
    2: 2,  # Channel 2 uses register 2
    3: 3   # Channel 3 uses register 3
}

# Define all supported operating modes with their min/max values
OPERATING_MODES = {
    "±5V": (-5.0, 5.0),    # Bipolar ±5V range
    "±10V": (-10.0, 10.0),  # Bipolar ±10V range 
    "0-5V": (0.0, 5.0),     # Unipolar 0-5V range
    "0-10V": (0.0, 10.0),   # Unipolar 0-10V range
    "4-20mA": (4.0, 20.0)   # Current mode (4mA = "zero")
}

class AdamController:
    """Class to handle all communication with the ADAM-6224 device"""
    
    def __init__(self):
        """Initialize controller with default values"""
        self.verify_active = False  # Flag for active verification
        self.client = None  # Will hold the Modbus client connection

    def connect(self):
        """Establish connection to ADAM device"""
        try:
            # Create new Modbus TCP client
            self.client = ModbusTcpClient(ADAM_IP, port=PORT)
            # Attempt connection and return status
            return self.client.connect()
        except Exception as e:
            # Log connection errors
            logger.error(f"Connection failed: {str(e)}")
            return False

    def disconnect(self):
        """Close the Modbus connection"""
        if self.client:  # If connection exists
            self.client.close()  # Close it
            self.client = None  # Clear reference

    def voltage_to_register(self, voltage, mode):
        """Convert voltage/mA value to 16-bit register value"""
        # Get min/max values for current mode
        min_val, max_val = OPERATING_MODES[mode]
        
        # Validate input is within range
        if not (min_val <= voltage <= max_val):
            raise ValueError(f"Value must be between {min_val} and {max_val}")
        
        # Scale value to 0-4095 range (12-bit resolution)
        return int(((voltage - min_val) / (max_val - min_val)) * 4095)

    def set_channel(self, channel, value, mode):
        """Set output value for specific channel"""
        # Validate channel number
        if not 0 <= channel <= 3:
            raise ValueError("Channel must be 0-3")
            
        # Get register address for this channel    
        reg_addr = AO_CHANNEL_REGISTERS[channel]
        # Convert physical value to register value
        reg_value = self.voltage_to_register(value, mode)
        
        # Check connection status
        if not self.client or not self.client.is_socket_open():
            if not self.connect():  # Reconnect if needed
                raise ConnectionError("Device connection failed")
        
        # Write value to register
        result = self.client.write_register(reg_addr, reg_value)
        # Return True if successful
        return not result.isError()

    def read_channel(self, channel):
        """Read current value from a channel"""
        # Check connection status
        if not self.client or not self.client.is_socket_open():
            if not self.connect():  # Reconnect if needed
                raise ConnectionError("Device connection failed")
                
        # Read register value
        response = self.client.read_holding_registers(
            AO_CHANNEL_REGISTERS[channel], 1)
            
        if response.isError():
            raise Exception("Read operation failed")
            
        return response.registers[0]  # Return raw register value

    def initialize_outputs(self):
        """Set all outputs to zero on startup"""
        try:
            if not self.connect():  # Connect to device
                return False
            
            # Set each channel to zero
            for ch in range(4):
                mode = "±5V"  # Default mode for initialization
                reset_value = 0.0  # Target value
                self.set_channel(ch, reset_value, mode)
            
            return True  # Success
        except Exception as e:
            logger.error(f"Initialization failed: {str(e)}")
            return False
        finally:
            self.disconnect()  # Always disconnect when done

    def shutdown_outputs(self):
        """Set all outputs to zero before closing"""
        try:
            if not self.connect():  # Connect to device
                return False
            
            # Reset each channel
            for ch in range(4):
                try:
                    mode = "±5V"  # Use voltage mode for shutdown
                    reset_value = 0.0  # Target value
                    self.set_channel(ch, reset_value, mode)
                    logger.info(f"Channel {ch} set to 0V during shutdown")
                except Exception as e:
                    logger.error(f"Failed to reset channel {ch}: {str(e)}")
            
            return True  # Success
        except Exception as e:
            logger.error(f"Shutdown failed: {str(e)}")
            return False
        finally:
            self.disconnect()  # Always disconnect when done

class Application:
    """Main application class for the GUI interface"""
    
    def __init__(self, root):
        """Initialize application with main window"""
        self.root = root  # Store main window
        self.controller = AdamController()  # Create device controller
        self.initialize_outputs()  # Set outputs to zero at start
        self.setup_ui()  # Build the user interface
        
        # Set handler for window close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
    def on_close(self):
        """Handle window closing event"""
        # Reset outputs before closing
        self.controller.shutdown_outputs()
        # Close the window
        self.root.destroy()
        
    def initialize_outputs(self):
        """Initialize hardware outputs to zero"""
        if self.controller.initialize_outputs():
            logger.info("All outputs initialized to zero")
        else:
            # Show warning if initialization failed
            messagebox.showwarning(
                "Initialization Warning", 
                "Could not initialize hardware outputs to zero")
        
    def setup_ui(self):
        """Build the user interface"""
        # Set window title
        self.root.title("ADAM-6224 Control Panel")
        
        # Create main frames for organization
        control_frame = tk.Frame(self.root, padx=10, pady=5)
        control_frame.pack()
        
        button_frame = tk.Frame(self.root, pady=10)
        button_frame.pack()
        
        output_frame = tk.Frame(self.root, padx=10, pady=5)
        output_frame.pack(fill=tk.BOTH, expand=True)
        
        # Lists to store channel controls
        self.entries = []  # Value entry fields
        self.mode_vars = []  # Mode selection variables
        self.range_labels = []  # Range display labels

        # Create controls for each channel
        for ch in range(4):
            # Frame for this channel's controls
            frame = tk.Frame(control_frame)
            frame.pack(fill=tk.X, pady=2)
            
            # Channel label
            tk.Label(frame, text=f"Channel {ch}:", width=10).pack(side=tk.LEFT)
            
            # Mode selection dropdown
            mode_var = tk.StringVar(value="±5V")  # Default mode
            mode_menu = ttk.Combobox(
                frame, 
                textvariable=mode_var, 
                values=list(OPERATING_MODES.keys()),
                state="readonly", 
                width=8)
            mode_menu.pack(side=tk.LEFT, padx=5)
            # Update range when mode changes
            mode_var.trace_add("write", lambda *_, ch=ch: self.update_range(ch))
            self.mode_vars.append(mode_var)  # Store reference
            
            # Value entry field
            entry = tk.Entry(frame, width=10)
            entry.insert(0, "0.0")  # Default value
            entry.pack(side=tk.LEFT, padx=5)
            self.entries.append(entry)  # Store reference
            
            # Range display label
            range_label = tk.Label(frame, text="Range: -5.0V to 5.0V")
            range_label.pack(side=tk.LEFT, padx=5)
            self.range_labels.append(range_label)  # Store reference
        
        # Verification checkbox frame
        verify_frame = tk.Frame(self.root)
        verify_frame.pack()
        
        # Verification checkbox
        self.verify_var = tk.BooleanVar()  # Stores checkbox state
        verify_check = tk.Checkbutton(
            verify_frame, 
            text="Real-time Verification",
            variable=self.verify_var, 
            command=self.toggle_verification)
        verify_check.pack()
        
        # Action buttons
        tk.Button(
            button_frame, 
            text="Apply Settings", 
            command=self.apply_settings).pack(side=tk.LEFT, padx=5)
            
        tk.Button(
            button_frame, 
            text="Reset All", 
            command=self.reset_all).pack(side=tk.LEFT, padx=5)
            
        tk.Button(
            button_frame, 
            text="Exit", 
            command=self.on_close).pack(side=tk.LEFT, padx=5)
        
        # Output console
        self.output_text = tk.Text(
            output_frame, 
            height=10, 
            width=80,
            wrap=tk.WORD, 
            state="normal")
        # Console scrollbar
        scrollbar = tk.Scrollbar(
            output_frame, 
            command=self.output_text.yview)
        
        # Configure scrolling
        self.output_text.config(yscrollcommand=scrollbar.set)
        self.output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Startup message
        self.log_message("Application started - all outputs initialized to zero")
    
    def update_range(self, channel):
        """Update range display when mode changes"""
        mode = self.mode_vars[channel].get()  # Get current mode
        min_val, max_val = OPERATING_MODES[mode]  # Get min/max values
        unit = "mA" if mode == "4-20mA" else "V"  # Determine unit
        # Update label text
        self.range_labels[channel].config(
            text=f"Range: {min_val:.1f}{unit} to {max_val:.1f}{unit}")
    
    def log_message(self, message):
        """Add message to output console"""
        self.output_text.insert(tk.END, message + "\n")  # Add message
        self.output_text.see(tk.END)  # Scroll to show new message
    
    def apply_settings(self):
        """Send current settings to device"""
        try:
            # Process each channel
            for ch in range(4):
                try:
                    # Get value from entry field
                    value = float(self.entries[ch].get())
                    # Get selected mode
                    mode = self.mode_vars[ch].get()
                    
                    # Set channel value
                    if self.controller.set_channel(ch, value, mode):
                        # Log success
                        self.log_message(
                            f"Channel {ch} set to {value:.2f} "
                            f"{'mA' if mode == '4-20mA' else 'V'} ({mode})")
                    else:
                        # Show error if failed
                        messagebox.showerror("Error", f"Failed to set channel {ch}")
                except ValueError as e:
                    messagebox.showwarning("Input Error", str(e))
                except ConnectionError as e:
                    messagebox.showerror("Connection Error", str(e))
                    break
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def reset_all(self):
        """Reset all channels to zero"""
        try:
            for ch in range(4):
                try:
                    # Update GUI
                    self.entries[ch].delete(0, tk.END)
                    self.entries[ch].insert(0, "0.0")
                    
                    # Get current mode
                    mode = self.mode_vars[ch].get()
                    
                    # Determine reset value (4mA for current mode)
                    reset_value = 4.0 if mode == "4-20mA" else 0.0
                    
                    # Update hardware
                    if self.controller.set_channel(ch, reset_value, mode):
                        # Log success
                        self.log_message(
                            f"Channel {ch} reset to {reset_value:.1f} "
                            f"{'mA' if mode == '4-20mA' else 'V'} ({mode})")
                    else:
                        self.log_message(f"Failed to reset channel {ch}")
                except Exception as e:
                    self.log_message(f"Channel {ch} reset error: {str(e)}")
        except Exception as e:
            messagebox.showerror("Reset Error", str(e))
    
    def toggle_verification(self):
        """Toggle real-time verification"""
        if self.verify_var.get():  # If checked
            # Start verification
            self.controller.verify_active = True
            # Run in background thread
            threading.Thread(
                target=self.verification_loop, 
                daemon=True).start()
            self.log_message("Started real-time verification")
        else:  # If unchecked
            # Stop verification
            self.controller.verify_active = False
            self.log_message("Stopped real-time verification")
    
    def verification_loop(self):
        """Check current values periodically"""
        while self.controller.verify_active:  # While active
            # Check each channel
            for ch in range(4):
                try:
                    # Get current mode
                    mode = self.mode_vars[ch].get()
                    # Read current value
                    reg_value = self.controller.read_channel(ch)
                    
                    # Convert to physical value
                    min_val, max_val = OPERATING_MODES[mode]
                    value = min_val + (reg_value / 4095) * (max_val - min_val)
                    
                    # Display current value
                    self.log_message(
                        f"Channel {ch} current: {value:.2f}"
                        f"{'mA' if mode == '4-20mA' else 'V'}")
                except Exception as e:
                    self.log_message(f"Verify channel {ch} error: {str(e)}")
            
            # Wait before next check
            time.sleep(VERIFICATION_INTERVAL)

# Main program entry point
if __name__ == "__main__":
    root = tk.Tk()  # Create main window
    app = Application(root)  # Create application
    root.mainloop()  # Start event loop