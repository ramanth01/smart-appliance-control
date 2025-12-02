import tkinter as tk
from tkinter import messagebox
import speech_recognition as sr
import pandas as pd
from datetime import datetime
import pyttsx3

# ================== HARDWARE LAYER (READY FOR FUTURE BOARD) ==================

HARDWARE_ENABLED = False  # Set to True later when you connect real board


def init_hardware():
    """
    Placeholder for hardware initialization.
    Later you can:
    - Open serial port (Arduino/ESP)
    - Setup GPIO pins (Raspberry Pi)
    """
    if HARDWARE_ENABLED:
        # Example (future):
        # import serial
        # global ser
        # ser = serial.Serial('COM3', 9600)
        pass


def send_to_hardware(appliance, action):
    """
    Placeholder for sending commands to hardware.
    appliance: 'Light', 'Fan', 'AC'
    action: 'ON' or 'OFF'
    """
    if HARDWARE_ENABLED:
        # Example protocol (future):
        # cmd = f"{appliance}:{action}\n"
        # ser.write(cmd.encode())
        pass


# ================== MAIN APPLICATION CLASS ==================


class SmartHomeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Appliances Control System - Smart + Hardware Ready")
        self.root.geometry("650x700")
        self.root.config(bg="#1E1E1E")

        # ---------- Speech Engine ----------
        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 160)

        # ---------- Logging ----------
        self.log_file_actions = "smart_appliance_log.xlsx"
        self.log_file_environment = "smart_environment_log.xlsx"

        # ---------- Appliance States ----------
        self.appliance_states = {
            "Light": "OFF",
            "Fan": "OFF",
            "AC": "OFF"
        }

        # Power ratings in Watts (simulated)
        self.appliance_power = {
            "Light": 10,
            "Fan": 60,
            "AC": 1500
        }

        # Energy usage in Wh (since app start)
        self.energy_wh = {
            "Light": 0.0,
            "Fan": 0.0,
            "AC": 0.0
        }

        # ---------- Simulation Variables ----------
        self.room_temp = 27.0  # starting temp
        self.last_update_time = datetime.now()
        self.last_env_log_time = datetime.now()

        # ---------- Schedules ----------
        self.schedules = []  # each: {'appliance', 'action', 'time', 'last_run_date'}

        # Smart mode flag
        self.smart_mode = tk.BooleanVar(value=False)

        # ---------- GUI ----------
        self.build_gui()

        # Initialize hardware (no effect now, ready for future)
        init_hardware()

        # ---------- Start periodic updates ----------
        self.update_loop()

    # ================== LOGGING ==================

    def log_action(self, appliance, action, method):
        df = pd.DataFrame([[datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            appliance, action, method]],
                          columns=["Date-Time", "Appliance", "Action", "Method"])
        try:
            old_df = pd.read_excel(self.log_file_actions)
            df = pd.concat([old_df, df], ignore_index=True)
        except FileNotFoundError:
            pass
        df.to_excel(self.log_file_actions, index=False)

    def log_environment(self):
        now = datetime.now()
        total_power = self.calculate_current_power()
        total_energy_kwh = sum(self.energy_wh.values()) / 1000.0

        df = pd.DataFrame([[
            now.strftime("%Y-%m-%d %H:%M:%S"),
            round(self.room_temp, 2),
            1 if self.appliance_states["Light"] == "ON" else 0,
            1 if self.appliance_states["Fan"] == "ON" else 0,
            1 if self.appliance_states["AC"] == "ON" else 0,
            round(total_power, 2),
            round(total_energy_kwh, 4)
        ]],
            columns=[
                "Date-Time",
                "Temperature(C)",
                "Light_ON",
                "Fan_ON",
                "AC_ON",
                "TotalPower(W)",
                "TotalEnergy(kWh)"
            ]
        )

        try:
            old_df = pd.read_excel(self.log_file_environment)
            df = pd.concat([old_df, df], ignore_index=True)
        except FileNotFoundError:
            pass
        df.to_excel(self.log_file_environment, index=False)

    # ================== APPLIANCE CONTROL ==================

    def update_status_label(self, appliance):
        state = self.appliance_states[appliance]
        label = self.status_labels[appliance]
        label.config(text=f"Status: {state}", fg="green" if state == "ON" else "red")

    def control_appliance(self, appliance, action, method="Manual"):
        self.appliance_states[appliance] = action
        self.update_status_label(appliance)
        self.log_action(appliance, action, method)
        send_to_hardware(appliance, action)
        messagebox.showinfo("Smart Control",
                            f"{appliance} turned {action} via {method}")

    # ================== COMMAND CONTROL ==================

    def process_command(self):
        cmd = self.command_entry.get().lower().strip()
        if not cmd:
            return

        if "light" in cmd:
            action = "ON" if "on" in cmd else "OFF"
            self.control_appliance("Light", action, method="Command")
        elif "fan" in cmd:
            action = "ON" if "on" in cmd else "OFF"
            self.control_appliance("Fan", action, method="Command")
        elif "ac" in cmd or "air conditioner" in cmd:
            action = "ON" if "on" in cmd else "OFF"
            self.control_appliance("AC", action, method="Command")
        else:
            messagebox.showwarning("Unknown Command", "Command not recognized.")

        self.command_entry.delete(0, tk.END)

    # ================== VOICE CONTROL ==================

    def voice_control(self):
        recognizer = sr.Recognizer()
        try:
            with sr.Microphone() as source:
                self.listening_label.config(text="🎤 Listening...", fg="red")
                self.root.update()
                self.engine.say("Listening now")
                self.engine.runAndWait()
                audio = recognizer.listen(source)
            self.listening_label.config(text="")
            cmd = recognizer.recognize_google(audio).lower()
            self.engine.say(f"You said {cmd}")
            self.engine.runAndWait()
            self.command_entry.delete(0, tk.END)
            self.command_entry.insert(0, cmd)
            self.process_command()
        except sr.UnknownValueError:
            self.engine.say("Sorry, I did not understand.")
            self.engine.runAndWait()
            messagebox.showerror("Error", "Could not understand your voice.")
            self.listening_label.config(text="")
        except sr.RequestError:
            self.engine.say("Speech recognition service error.")
            self.engine.runAndWait()
            messagebox.showerror("Error", "Speech recognition service unavailable.")
            self.listening_label.config(text="")

    # ================== SCHEDULING ==================

    def add_schedule(self):
        appliance = self.schedule_appliance_var.get()
        action = self.schedule_action_var.get()
        time_str = self.schedule_time_entry.get().strip()

        if not time_str:
            messagebox.showwarning("Schedule", "Please enter time in HH:MM format.")
            return

        try:
            datetime.strptime(time_str, "%H:%M")
        except ValueError:
            messagebox.showerror("Invalid Time", "Please enter time in HH:MM format (24-hour).")
            return

        schedule = {
            "appliance": appliance,
            "action": action,
            "time": time_str,
            "last_run_date": None
        }
        self.schedules.append(schedule)
        self.refresh_schedule_list()
        messagebox.showinfo("Schedule", f"Scheduled {appliance} {action} at {time_str} daily.")

    def refresh_schedule_list(self):
        self.schedule_listbox.delete(0, tk.END)
        for sch in self.schedules:
            text = f"{sch['time']} - {sch['appliance']} -> {sch['action']}"
            self.schedule_listbox.insert(tk.END, text)

    # ================== SIMULATION & SMART LOGIC ==================

    def calculate_current_power(self):
        power = 0.0
        for appliance, state in self.appliance_states.items():
            if state == "ON":
                power += self.appliance_power[appliance]
        return power

    def update_simulation(self, delta_seconds):
        # Temperature behavior
        if self.appliance_states["AC"] == "ON":
            # Cool down faster
            self.room_temp -= 0.05 * delta_seconds
            if self.room_temp < 22.0:
                self.room_temp = 22.0
        else:
            # Warm up slowly towards 32
            self.room_temp += 0.02 * delta_seconds
            if self.room_temp > 32.0:
                self.room_temp = 32.0

        # Energy usage
        for appliance, state in self.appliance_states.items():
            if state == "ON":
                self.energy_wh[appliance] += self.appliance_power[appliance] * (delta_seconds / 3600.0)

        # Update labels
        self.temp_label.config(text=f"Room Temp: {self.room_temp:.1f} °C")
        total_energy_kwh = sum(self.energy_wh.values()) / 1000.0
        self.energy_label.config(
            text=f"Total Energy Used: {total_energy_kwh:.3f} kWh"
        )
        current_power = self.calculate_current_power()
        self.power_label.config(text=f"Current Power: {current_power:.1f} W")

    def apply_smart_mode(self):
        if not self.smart_mode.get():
            return

        # Smart rule 1: AC auto control based on temperature
        if self.room_temp > 28.0 and self.appliance_states["AC"] == "OFF":
            self.control_appliance("AC", "ON", method="Smart Mode")
        elif self.room_temp < 24.0 and self.appliance_states["AC"] == "ON":
            self.control_appliance("AC", "OFF", method="Smart Mode")

        # Smart rule 2: Fan assists cooling when hot
        if self.room_temp > 29.0 and self.appliance_states["Fan"] == "OFF":
            self.control_appliance("Fan", "ON", method="Smart Mode")
        elif self.room_temp < 25.0 and self.appliance_states["Fan"] == "ON" and self.appliance_states["AC"] == "OFF":
            self.control_appliance("Fan", "OFF", method="Smart Mode")

        # Smart rule 3: Light auto ON in evening & OFF late night (simulated via real time)
        now = datetime.now()
        hour = now.hour
        if 18 <= hour <= 23 and self.appliance_states["Light"] == "OFF":
            self.control_appliance("Light", "ON", method="Smart Mode")
        if (hour >= 0 and hour <= 5) and self.appliance_states["Light"] == "ON":
            self.control_appliance("Light", "OFF", method="Smart Mode")

    def check_schedules(self):
        now = datetime.now()
        now_time = now.strftime("%H:%M")
        today_str = now.strftime("%Y-%m-%d")

        for sch in self.schedules:
            if sch["time"] == now_time and sch.get("last_run_date") != today_str:
                self.control_appliance(
                    sch["appliance"],
                    sch["action"],
                    method="Schedule"
                )
                sch["last_run_date"] = today_str

    def update_loop(self):
        now = datetime.now()
        delta_seconds = (now - self.last_update_time).total_seconds()
        if delta_seconds <= 0:
            delta_seconds = 1
        self.last_update_time = now

        # Simulation
        self.update_simulation(delta_seconds)

        # Smart automation
        self.apply_smart_mode()

        # Schedules
        self.check_schedules()

        # Periodic environment logging (every 30 seconds)
        if (now - self.last_env_log_time).total_seconds() >= 30:
            self.log_environment()
            self.last_env_log_time = now

        # Schedule next tick
        self.root.after(1000, self.update_loop)

    # ================== GUI BUILDING ==================

    def build_gui(self):
        title = tk.Label(
            self.root,
            text="Smart Appliances Control - Smart System",
            font=("Arial", 18, "bold"),
            bg="#1E1E1E",
            fg="white"
        )
        title.pack(pady=10)

        # Appliance frames
        self.status_labels = {}

        container = tk.Frame(self.root, bg="#1E1E1E")
        container.pack(pady=5)

        # Light
        light_frame = tk.LabelFrame(container, text="Light", bg="#1E1E1E", fg="white", font=("Arial", 12))
        light_frame.grid(row=0, column=0, padx=10, pady=10)
        light_status = tk.Label(light_frame, text="Status: OFF", fg="red", bg="#1E1E1E", font=("Arial", 11))
        light_status.pack()
        tk.Button(light_frame, text="Turn ON", width=10,
                  command=lambda: self.control_appliance("Light", "ON")).pack(side="left", padx=5, pady=5)
        tk.Button(light_frame, text="Turn OFF", width=10,
                  command=lambda: self.control_appliance("Light", "OFF")).pack(side="left", padx=5, pady=5)
        self.status_labels["Light"] = light_status

        # Fan
        fan_frame = tk.LabelFrame(container, text="Fan", bg="#1E1E1E", fg="white", font=("Arial", 12))
        fan_frame.grid(row=0, column=1, padx=10, pady=10)
        fan_status = tk.Label(fan_frame, text="Status: OFF", fg="red", bg="#1E1E1E", font=("Arial", 11))
        fan_status.pack()
        tk.Button(fan_frame, text="Turn ON", width=10,
                  command=lambda: self.control_appliance("Fan", "ON")).pack(side="left", padx=5, pady=5)
        tk.Button(fan_frame, text="Turn OFF", width=10,
                  command=lambda: self.control_appliance("Fan", "OFF")).pack(side="left", padx=5, pady=5)
        self.status_labels["Fan"] = fan_status

        # AC
        ac_frame = tk.LabelFrame(container, text="AC", bg="#1E1E1E", fg="white", font=("Arial", 12))
        ac_frame.grid(row=0, column=2, padx=10, pady=10)
        ac_status = tk.Label(ac_frame, text="Status: OFF", fg="red", bg="#1E1E1E", font=("Arial", 11))
        ac_status.pack()
        tk.Button(ac_frame, text="Turn ON", width=10,
                  command=lambda: self.control_appliance("AC", "ON")).pack(side="left", padx=5, pady=5)
        tk.Button(ac_frame, text="Turn OFF", width=10,
                  command=lambda: self.control_appliance("AC", "OFF")).pack(side="left", padx=5, pady=5)
        self.status_labels["AC"] = ac_status

        # Command Control
        cmd_label = tk.Label(self.root, text="Type Command (e.g., 'turn on light'):",
                             bg="#1E1E1E", fg="white", font=("Arial", 12))
        cmd_label.pack(pady=5)
        self.command_entry = tk.Entry(self.root, width=50)
        self.command_entry.pack(pady=5)
        tk.Button(self.root, text="Execute Command",
                  command=self.process_command, width=20).pack(pady=5)

        # Voice Control
        self.listening_label = tk.Label(self.root, text="", bg="#1E1E1E", fg="red", font=("Arial", 12, "bold"))
        self.listening_label.pack(pady=5)
        tk.Button(self.root, text="🎤 Voice Control", command=self.voice_control,
                  bg="#4CAF50", fg="white", width=20).pack(pady=10)

        # Simulation Info
        sim_frame = tk.LabelFrame(self.root, text="Environment & Energy", bg="#1E1E1E", fg="white",
                                  font=("Arial", 12))
        sim_frame.pack(pady=10, fill="x", padx=10)
        self.temp_label = tk.Label(sim_frame, text=f"Room Temp: {self.room_temp:.1f} °C",
                                   bg="#1E1E1E", fg="white", font=("Arial", 11))
        self.temp_label.pack(anchor="w", padx=10, pady=2)
        self.power_label = tk.Label(sim_frame, text="Current Power: 0.0 W",
                                    bg="#1E1E1E", fg="white", font=("Arial", 11))
        self.power_label.pack(anchor="w", padx=10, pady=2)
        self.energy_label = tk.Label(sim_frame, text="Total Energy Used: 0.000 kWh",
                                     bg="#1E1E1E", fg="white", font=("Arial", 11))
        self.energy_label.pack(anchor="w", padx=10, pady=2)

        # Smart Mode
        smart_frame = tk.LabelFrame(self.root, text="Smart Mode", bg="#1E1E1E", fg="white",
                                    font=("Arial", 12))
        smart_frame.pack(pady=10, fill="x", padx=10)
        tk.Checkbutton(smart_frame, text="Enable Smart Mode (Auto AC/Fan/Light)",
                       variable=self.smart_mode, bg="#1E1E1E", fg="white",
                       selectcolor="#333333", font=("Arial", 11)).pack(anchor="w", padx=10, pady=5)

        # Scheduling
        schedule_frame = tk.LabelFrame(self.root, text="Scheduling (Daily)", bg="#1E1E1E", fg="white",
                                       font=("Arial", 12))
        schedule_frame.pack(pady=10, fill="x", padx=10)

        # Controls for schedule input
        input_frame = tk.Frame(schedule_frame, bg="#1E1E1E")
        input_frame.pack(anchor="w", padx=10, pady=5)

        tk.Label(input_frame, text="Appliance:", bg="#1E1E1E", fg="white").grid(row=0, column=0, padx=5)
        self.schedule_appliance_var = tk.StringVar(value="Light")
        tk.OptionMenu(input_frame, self.schedule_appliance_var, "Light", "Fan", "AC").grid(row=0, column=1, padx=5)

        tk.Label(input_frame, text="Action:", bg="#1E1E1E", fg="white").grid(row=0, column=2, padx=5)
        self.schedule_action_var = tk.StringVar(value="ON")
        tk.OptionMenu(input_frame, self.schedule_action_var, "ON", "OFF").grid(row=0, column=3, padx=5)

        tk.Label(input_frame, text="Time (HH:MM 24h):", bg="#1E1E1E", fg="white").grid(row=0, column=4, padx=5)
        self.schedule_time_entry = tk.Entry(input_frame, width=8)
        self.schedule_time_entry.grid(row=0, column=5, padx=5)

        tk.Button(input_frame, text="Add Schedule", command=self.add_schedule).grid(row=0, column=6, padx=5)

        # List of schedules
        self.schedule_listbox = tk.Listbox(schedule_frame, width=70, height=5)
        self.schedule_listbox.pack(padx=10, pady=5)


if __name__ == "__main__":
    root = tk.Tk()
    app = SmartHomeApp(root)
    root.mainloop()
