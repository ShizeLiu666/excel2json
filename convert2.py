import sys
import json
import pandas as pd
import re
import io

def extract_text_from_sheet(sheet_df):
    text_list = []
    for value in sheet_df.values.flatten():
        if pd.notna(value) and isinstance(value, str):
            value = value.replace('\uff08', '(').replace('\uff09', ')').replace('\uff1a', ':')
            value = re.sub(r'\(.*?\)', '', value)
            text_list.extend([text.strip() for text in value.split('\n') if text.strip()])
    return text_list

def process_excel_to_json(file_content):
    xl = pd.ExcelFile(file_content)
    all_text_data = {}
    for sheet_name in xl.sheet_names:
        if "Programming Details" in sheet_name:
            df = xl.parse(sheet_name)
            all_text_data["programming details"] = extract_text_from_sheet(df)
    
    return all_text_data if all_text_data else None

DevicesInSceneControl = {
    "Dimmer Type": [
        "KBSKTDIM", "D300IB", "D300IB2", "DH10VIB", 
        "DM300BH", "D0-10IB", "DDAL"
    ],
    "Relay Type": [
        "KBSKTREL", "S2400IB2", "RM1440BH", "KBSKTR", "Z2"
    ],
    "Curtain Type": [
        "C300IBH"
    ],
    "Fan Type": [
        "FC150A2"
    ],
    "RGB Type": [
        "KB8RGBG", "KB36RGBS", "KB9TWG", "KB12RGBD", 
        "KB12RGBG"
    ],
    "PowerPoint Type": {
        "Single-Way": [
            "H1PPWVBX"
        ],
        "Two-Way": [
            "K2PPHB", "H2PPHB", "H2PPWHB"
        ]
    }
}

device_name_to_type = {}

def reset_device_name_to_type():
    global device_name_to_type
    device_name_to_type = {}

def process_devices(split_data):
    devices_content = split_data.get("devices", [])
    devices_data = []
    current_shortname = None

    global device_name_to_type

    for line in devices_content:
        line = line.strip()

        if line.startswith("NAME:"):
            current_shortname = line.replace("NAME:", "").strip()
            continue
        
        if line.startswith("QTY:"):
            continue

        device_type = None
        for dtype, models in DevicesInSceneControl.items():
            if isinstance(models, dict):
                for sub_type, sub_models in models.items():
                    for model in sub_models:
                        if model in current_shortname or current_shortname in model:
                            device_type = f"{dtype} ({sub_type})"
                            break
                    if device_type:
                        break
            else:
                for model in models:
                    if model in current_shortname or current_shortname in model:
                        device_type = dtype
                        break
            if device_type:
                break

        if current_shortname:
            device_info = {
                "appearanceShortname": current_shortname,
                "deviceName": line
            }
            if device_type:
                device_info["deviceType"] = device_type
                device_name_to_type[line] = device_type
            devices_data.append(device_info)

    return {"devices": devices_data}

def process_groups(split_data):
    groups_content = split_data.get("groups", [])
    groups_data = []
    current_group = None

    for line in groups_content:
        line = line.strip()

        if line.startswith("NAME:"):
            current_group = line.replace("NAME:", "").strip()
            continue
        
        if line.startswith("DEVICE CONTROL:"):
            continue

        if current_group:
            groups_data.append({
                "groupName": current_group,
                "devices": line
            })

    return {"groups": groups_data}

scene_output_templates = {
    "Relay Type": lambda name, status: {
        "name": name,
        "status": status,
        "statusConditions": {}
    },
    "Curtain Type": lambda name, status: {
        "name": name,
        "status": status,
        "statusConditions": {
            "position": 100 if status == "OPEN" else 0
        }
    },
    "Dimmer Type": lambda name, status, level=100: {
        "name": name,
        "status": status,
        "statusConditions": {
            "level": level 
        }
    },
    "Fan Type": lambda name, status, relay_status, speed: {
        "name": name,
        "status": status,
        "statusConditions": {
            "relay": relay_status,
            "speed": speed
        }
    },
    "PowerPoint Type": {
        "Two-Way": lambda name, left_power, right_power: {
            "name": name,
            "statusConditions": {
                "leftPowerOnOff": left_power,
                "rightPowerOnOff": right_power
            }
        },
        "Single-Way": lambda name, power: {
            "name": name,
            "statusConditions": {
                "rightPowerOnOff": power
            }
        }
    }
}

def handle_fan_type(parts):
    device_name = parts[0]
    status = parts[1]
    relay_status = parts[3]
    speed = int(parts[5])
    return [scene_output_templates["Fan Type"](device_name, status, relay_status, speed)]

def handle_dimmer_type(parts):
    contents = []
    status_index = next(i for i, part in enumerate(parts) if part in ["ON", "OFF"])
    status = parts[status_index]

    level = 100
    
    if status == "ON" and len(parts) > status_index + 1:
        try:
            level_part = parts[status_index + 1].replace("+", "").replace("%", "").strip()
            level = int(level_part)
        except ValueError:
            level = 100
    elif status == "OFF":
        level = 0

    for entry in parts[:status_index]: 
        device_name = entry.strip().strip(",")  
        contents.append(scene_output_templates["Dimmer Type"](device_name, status, level))

    return contents

def handle_relay_type(parts):
    contents = []
    status = parts[-1]

    for entry in parts[:-1]:
        device_name = entry.strip().strip(",")  
        contents.append(scene_output_templates["Relay Type"](device_name, status))

    return contents

def handle_curtain_type(parts):
    contents = []
    status = parts[-1]

    for entry in parts[:-1]:  
        device_name = entry.strip().strip(",")  
        contents.append(scene_output_templates["Curtain Type"](device_name, status))

    return contents

def handle_powerpoint_type(parts, device_type):
    contents = []

    if "Two-Way" in device_type:
        right_power = parts[-1]
        left_power = parts[-2]
        device_names = parts[:-2]

        for device_name in device_names:
            device_name = device_name.strip().strip(",") 
            contents.append(scene_output_templates["PowerPoint Type"]["Two-Way"](device_name, left_power, right_power))

    elif "Single-Way" in device_type:
        power = parts[-1]
        device_names = parts[:-1]

        for device_name in device_names:
            device_name = device_name.strip().strip(",")
            contents.append(scene_output_templates["PowerPoint Type"]["Single-Way"](device_name, power))

    return contents

def determine_device_type(device_name):
    original_device_name = device_name.strip().strip(',')
    
    if not original_device_name:
        print(f"Error: Detected empty or invalid device name: '{original_device_name}'")
        raise ValueError("设备名称不能为空。")

    device_type = device_name_to_type.get(original_device_name)
    
    if device_type:
        return device_type
    else:
        raise ValueError(f"无法确定设备类型：'{original_device_name}'")

def parse_scene_content(scene_name, content_lines):
    contents = []
    
    for line in content_lines:
        parts = line.split()
        if len(parts) < 2:
            continue
        try:
            device_type = determine_device_type(parts[0]) 
        except ValueError:
            continue
        
        if device_type == "Fan Type":
            contents.extend(handle_fan_type(parts))
        elif device_type == "Relay Type":
            contents.extend(handle_relay_type(parts))
        elif device_type == "Curtain Type":
            contents.extend(handle_curtain_type(parts))
        elif device_type == "Dimmer Type":
            contents.extend(handle_dimmer_type(parts))
        elif "PowerPoint Type" in device_type:
            if "Two-Way" in device_type:
                contents.extend(handle_powerpoint_type(parts, "Two-Way PowerPoint Type"))
            elif "Single-Way" in device_type:
                contents.extend(handle_powerpoint_type(parts, "Single-Way PowerPoint Type"))
    
    return contents

def process_scenes(split_data):
    scenes_content = split_data.get("scenes", [])
    scenes_data = {}
    current_scene = None

    for i, line in enumerate(scenes_content):
        line = line.strip()
        
        if line.startswith("CONTROL CONTENT:"):
            continue
        
        if line.startswith("NAME:"):
            if current_scene and current_scene in scenes_data:
                scenes_data[current_scene] = scenes_data[current_scene]

            current_scene = line.replace("NAME:", "").strip()
            if current_scene not in scenes_data:
                scenes_data[current_scene] = []
        elif current_scene:
            try:
                scenes_data[current_scene].extend(parse_scene_content(current_scene, [line]))
            except ValueError:
                continue

    scenes_output = [{"sceneName": scene_name, "contents": contents} for scene_name, contents in scenes_data.items()]

    return {"scenes": scenes_output}

def process_remote_controls(split_data):
    remote_controls_content = split_data.get("remoteControls", [])
    remote_controls_data = []
    current_remote = None
    current_links = []

    for line in remote_controls_content:
        line = line.strip()

        if line.startswith("TOTAL"):
            continue

        if line.startswith("NAME:"):
            if current_remote:
                remote_controls_data.append({
                    "remoteName": current_remote,
                    "links": current_links
                })
            current_remote = line.replace("NAME:", "").strip()
            current_links = []
        
        elif line.startswith("LINK:"):
            continue

        else:
            parts = line.split(":")
            if len(parts) < 2:
                continue

            link_index = int(parts[0].strip()) - 1
            link_description = parts[1].strip()

            action = "NORMAL"
            if " - " in link_description:
                link_description, action = link_description.rsplit(" - ", 1)
                action = action.strip().upper()

            if link_description.startswith("SCENE"):
                link_type = 2
                link_name = link_description.replace("SCENE", "").strip()
            elif link_description.startswith("GROUP"):
                link_type = 1
                link_name = link_description.replace("GROUP", "").strip()
            elif link_description.startswith("DEVICE"):
                link_type = 0
                link_name = link_description.replace("DEVICE", "").strip()

            current_links.append({
                "linkIndex": link_index,
                "linkType": link_type,
                "linkName": link_name,
                "action": action
            })

    if current_remote:
        remote_controls_data.append({
            "remoteName": current_remote,
            "links": current_links
        })

    return {"remoteControls": remote_controls_data}

def split_json_file(input_data):
    content = input_data.get("programming details", [])
    split_keywords = {
        "devices": "KASTA DEVICE",
        "groups": "KASTA GROUP",
        "scenes": "KASTA SCENE",
        "remoteControls": "REMOTE CONTROL LINK"
    }
    split_data = {
        "devices": [],
        "groups": [],
        "scenes": [],
        "remoteControls": []
    }
    current_key = None
    for line in content:
        if line in split_keywords.values():
            current_key = next(key for key, value in split_keywords.items() if value == line)
            continue
        if current_key:
            split_data[current_key].append(line)
    
    result = {}
    result.update(process_devices(split_data))
    result.update(process_groups(split_data))
    result.update(process_scenes(split_data))
    result.update(process_remote_controls(split_data))
    return result

def main():
    try:
        file_content = sys.stdin.buffer.read()
        all_text_data = process_excel_to_json(io.BytesIO(file_content))
        if all_text_data:
            result = split_json_file(all_text_data)
            json_output = json.dumps(result)
            print(json_output)
        else:
            print(json.dumps({"error": "No matching worksheets found"}))
    except Exception as e:
        error_message = f"Error: {e}"
        print(json.dumps({"error": error_message}), file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
