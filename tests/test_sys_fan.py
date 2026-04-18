"""
Test script for reading SYS FAN RPM (excluding CPU and PUMP) via LibreHardwareMonitor.
Tests PWM control functionality for each detected fan.

Usage:
    python tests/test_sys_fan.py
    (Run as Administrator for full functionality)
"""

import os
import sys
import time
import threading
import logging

# Setup logging
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)


class SysFanTester:
    """Test and monitor SYS FAN controls via LibreHardwareMonitor."""
    
    def __init__(self):
        self.computer = None
        self.lock = threading.Lock()
        self.fans = []  # List of (control, sensor_name, subfan_name)
        self.rpm_sensors = {}  # sensor_name -> RPM sensor object
    
    def init_lhm(self):
        """Initialize LibreHardwareMonitor."""
        try:
            import clr
            # Go up one level from tests/ to project root, then into lib/lhm
            dll_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "lib", "lhm")
            if not os.path.exists(dll_dir):
                log.error(f"LibreHardwareMonitor DLLs not found at {dll_dir}")
                log.error("Please ensure the 'lib/lhm' directory exists with the DLL files.")
                return False
            
            log.info(f"Loading LibreHardwareMonitor from: {dll_dir}")
            clr.AddReference(os.path.join(dll_dir, "LibreHardwareMonitorLib.dll"))
            from LibreHardwareMonitor.Hardware import Computer
            
            self.computer = Computer()
            self.computer.IsCpuEnabled = True
            self.computer.IsGpuEnabled = False
            self.computer.IsMemoryEnabled = True
            self.computer.IsMotherboardEnabled = True
            self.computer.IsControllerEnabled = True
            self.computer.IsNetworkEnabled = False
            self.computer.IsStorageEnabled = False
            self.computer.Open()
            
            log.info("LibreHardwareMonitor initialized successfully")
            return True
            
        except Exception as e:
            log.error(f"Failed to initialize LibreHardwareMonitor: {e}")
            return False
    
    def discover_sys_fans(self):
        """
        Discover all SYS FAN controls on the motherboard.
        Excludes CPU and PUMP fans (RPM >= 3000).
        """
        if not self.computer:
            log.error("LibreHardwareMonitor not initialized")
            return False
        
        with self.lock:
            try:
                motherboard = None
                for hw in self.computer.Hardware:
                    if hw.HardwareType.ToString() == "Motherboard":
                        motherboard = hw
                        break
                
                if not motherboard:
                    log.error("No motherboard found")
                    return False
                
                log.info(f"Found motherboard: {motherboard.Name}")
                
                # First pass: collect all RPM sensors for later reference
                for sub_hw in motherboard.SubHardware:
                    try:
                        sub_hw.Update()
                    except Exception:
                        pass
                    
                    for sensor in sub_hw.Sensors:
                        if sensor.SensorType.ToString() == "Fan":
                            self.rpm_sensors[sensor.Name] = sensor
                
                # Second pass: discover fan controls
                control_count = 0
                for sub_hw in motherboard.SubHardware:
                    sub_hw_name = sub_hw.Name
                    
                    try:
                        sub_hw.Update()
                    except Exception:
                        pass
                    
                    # Get RPM sensors for this sub-hardware
                    rpm_list = [s for s in sub_hw.Sensors if s.SensorType.ToString() == "Fan"]
                    
                    # Process control sensors
                    control_idx = 0
                    for sensor in sub_hw.Sensors:
                        if sensor.SensorType.ToString() != "Control":
                            continue
                        
                        ctrl = sensor.Control
                        if ctrl is None:
                            continue
                        
                        # Get corresponding RPM sensor
                        rpm_sensor = rpm_list[control_idx] if control_idx < len(rpm_list) else None
                        rpm_val = None
                        
                        if rpm_sensor:
                            try:
                                v = rpm_sensor.Value
                                rpm_val = float(v) if v is not None else None
                            except Exception:
                                pass
                        
                        # Skip disconnected fans (RPM = 0)
                        if rpm_val == 0:
                            log.info(
                                f"  Skipping {sensor.Name} on {sub_hw_name} (disconnected, RPM=0)"
                            )
                            control_idx += 1
                            continue
                        
                        # Skip pumps (RPM >= 3000)
                        if rpm_val is not None and rpm_val >= 3000:
                            log.info(
                                f"  Skipping {sensor.Name} on {sub_hw_name} "
                                f"(pump/high-RPM, RPM={rpm_val:.0f})"
                            )
                            control_idx += 1
                            continue
                        
                        # This is a valid SYS FAN
                        fan_id = len(self.fans)
                        mode = ctrl.ControlMode.ToString()
                        rpm_hint = f"RPM={rpm_val:.0f}" if rpm_val else "RPM=unknown"
                        
                        log.info(
                            f"  [{fan_id}] {sensor.Name} on {sub_hw_name} "
                            f"| Mode={mode} | {rpm_hint}"
                        )
                        
                        self.fans.append((ctrl, sensor.Name, sub_hw_name, rpm_sensor))
                        control_idx += 1
                
                if self.fans:
                    log.info(f"\nDiscovered {len(self.fans)} SYS FAN control(s)")
                    return True
                else:
                    log.warning("No SYS FAN controls found")
                    return False
            
            except Exception as e:
                log.error(f"Error discovering fans: {e}")
                return False
    
    def read_rpm_all(self):
        """Read current RPM for all discovered fans."""
        if not self.fans:
            log.warning("No fans discovered")
            return None
        
        with self.lock:
            rpms = {}
            try:
                for ctrl, ctrl_name, sub_hw_name, rpm_sensor in self.fans:
                    rpm_val = None
                    if rpm_sensor:
                        try:
                            v = rpm_sensor.Value
                            rpm_val = float(v) if v is not None else None
                        except Exception as e:
                            log.debug(f"Error reading RPM for {ctrl_name}: {e}")
                    
                    rpms[ctrl_name] = (sub_hw_name, rpm_val)
                
                return rpms
            except Exception as e:
                log.error(f"Error reading RPM: {e}")
                return None
    
    def test_pwm_control(self, fan_idx: int, pwm_values: list = None):
        """
        Test PWM control for a specific fan.
        """
        if fan_idx < 0 or fan_idx >= len(self.fans):
            log.error(f"Invalid fan index: {fan_idx}")
            return False
        
        ctrl, ctrl_name, sub_hw_name, rpm_sensor = self.fans[fan_idx]
        
        if pwm_values is None:
            pwm_values = [0, 30, 50, 70, 100]
        
        log.info(f"\nTesting PWM control for fan[{fan_idx}]: {ctrl_name}")
        log.info(f"  Location: {sub_hw_name}")
        log.info(f"  PWM steps: {pwm_values}")
        
        initial_rpm = None
        if rpm_sensor:
            try:
                v = rpm_sensor.Value
                initial_rpm = float(v) if v is not None else None
            except Exception:
                pass
        
        log.info(f"  Initial RPM: {initial_rpm}")
        log.info(f"  Initial ControlMode: {ctrl.ControlMode.ToString()}")
        
        results = []
        for pwm in pwm_values:
            try:
                ctrl.SetSoftware(pwm)
                time.sleep(0.5)
                
                new_software_val = ctrl.SoftwareValue
                new_mode = ctrl.ControlMode.ToString()
                
                new_rpm = None
                if rpm_sensor:
                    try:
                        v = rpm_sensor.Value
                        new_rpm = float(v) if v is not None else None
                    except Exception:
                        pass
                
                result = {
                    "pwm_set": pwm,
                    "software_value": new_software_val,
                    "control_mode": new_mode,
                    "rpm": new_rpm,
                }
                results.append(result)
                
                rpm_status = "changed" if (initial_rpm and new_rpm and new_rpm != initial_rpm) else "unchanged"
                log.info(
                    f"  PWM={pwm:3d} | SoftwareValue={new_software_val} | "
                    f"Mode={new_mode} | RPM={new_rpm} ({rpm_status})"
                )
                
            except Exception as e:
                log.error(f"  PWM={pwm} | Error: {e}")
                results.append({"pwm_set": pwm, "error": str(e)})
        
        # Reset to default
        try:
            ctrl.SetDefault()
            time.sleep(0.5)
            log.info("  Reset to default control mode")
        except Exception as e:
            log.error(f"  Failed to reset: {e}")
        
        return results
    
    def interactive_test(self):
        """Run interactive test session."""
        log.info("=" * 60)
        log.info("SYS FAN Interactive Test")
        log.info("=" * 60)
        
        # Initialize
        if not self.init_lhm():
            return False
        
        # Discover fans
        if not self.discover_sys_fans():
            return False
        
        # Interactive loop
        while True:
            print()
            print("Options:")
            print("  1. Read all RPMs")
            print("  2. Test PWM control")
            print("  3. Test all fans PWM")
            print("  4. Exit")
            
            choice = input("\nSelect option (1-4): ").strip()
            
            if choice == "1":
                rpms = self.read_rpm_all()
                if rpms:
                    for ctrl_name, (sub_hw_name, rpm_val) in rpms.items():
                        status = f"{rpm_val:.0f} RPM" if rpm_val is not None else "N/A"
                        print(f"  {ctrl_name} ({sub_hw_name}): {status}")
                else:
                    print("  No RPM data available")
            
            elif choice == "2":
                if not self.fans:
                    print("  No fans discovered")
                    continue
                
                print("\nAvailable fans:")
                for i, (ctrl, ctrl_name, sub_hw_name, _) in enumerate(self.fans):
                    print(f"  [{i}] {ctrl_name} ({sub_hw_name})")
                
                try:
                    fan_idx = int(input("\nSelect fan index: ").strip())
                    pwm_input = input("PWM values (comma-separated, default 0,30,50,70,100): ").strip()
                    if pwm_input:
                        pwm_values = [int(x.strip()) for x in pwm_input.split(",")]
                    else:
                        pwm_values = [0, 30, 50, 70, 100]
                    
                    self.test_pwm_control(fan_idx, pwm_values)
                except ValueError:
                    print("  Invalid input")
                except Exception as e:
                    print(f"  Error: {e}")
            
            elif choice == "3":
                for i in range(len(self.fans)):
                    self.test_pwm_control(i)
            
            elif choice == "4":
                if self.computer:
                    self.computer.Close()
                log.info("Exiting test")
                break
            
            else:
                print("  Invalid option")


def main():
    tester = SysFanTester()
    tester.interactive_test()


if __name__ == "__main__":
    main()
