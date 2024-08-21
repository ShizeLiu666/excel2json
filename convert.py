import sys
import openpyxl
import pandas as pd
import json
import io

print("Using openpyxl version:", openpyxl.__version__, file=sys.stderr)
print("Using pandas version:", pd.__version__, file=sys.stderr)

def extract_text_from_sheet(sheet_df):
    text_list = []
    for value in sheet_df.values.flatten():
        if pd.notna(value) and isinstance(value, str):
            value = value.replace('\uff08', '(').replace('\uff09', ')').replace('\uff1a', ':')
            value = value.replace("AK", "").replace("ES", "")
            value = re.sub(r'\(.*?\)', '', value)  # 使用正则表达式去除括号中的内容
            text_list.extend([text.strip() for text in value.split('\n') if text.strip()])
    return text_list

def process_excel_to_json(file_content):
    try:
        xl = pd.ExcelFile(io.BytesIO(file_content), engine='openpyxl')  # 使用openpyxl引擎
        all_text_data = {}
        for sheet_name in xl.sheet_names:
            if "Programming Details" in sheet_name: 
                df = xl.parse(sheet_name, engine='openpyxl')  # 确保使用openpyxl引擎
                all_text_data["programming details"] = extract_text_from_sheet(df)
        if not all_text_data:
            return None
        return all_text_data
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return None

def process_devices(split_data):
    devices_content = split_data.get("devices", [])
    devices_data = []
    current_shortname = None

    for line in devices_content:
        line = line.strip()

        if line.startswith("NAME:"):
            current_shortname = line.replace("NAME:", "").strip()
            continue
        
        if line.startswith("QTY:"):
            continue

        if current_shortname:
            devices_data.append({
                "appearanceShortname": current_shortname,
                "deviceName": line
            })

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

def parse_scene_content(scene_name, content_lines):
    contents = []
    for line in content_lines:
        parts = line.split()
        if len(parts) < 2:
            continue

        status = parts[-1] if parts[-1] in ["ON", "OFF"] else parts[-2]
        level = 100 if status == "ON" else 0
        
        if '+' in parts[-1]:
            try:
                level = int(parts[-1].replace("%", "").replace("+", ""))
            except (ValueError, IndexError):
                pass
        
        device_names = " ".join(parts[:-1] if parts[-1] in ["ON", "OFF"] else parts[:-2]).split(",")

        for name in device_names:
            contents.append({
                "name": name.strip(),
                "status": status,
                "statusConditions": {
                    "level": level
                }
            })
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
            current_scene = line.replace("NAME:", "").strip()
            if current_scene not in scenes_data:
                scenes_data[current_scene] = []
        elif current_scene:
            scenes_data[current_scene].extend(parse_scene_content(current_scene, [line]))

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
            else:
                continue

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
        all_text_data = process_excel_to_json(file_content)
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