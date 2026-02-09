import re
import PyPDF2

def extract_registered_architect_from_bytes(pdf_bytesio):
    try:
        data = {
            "status": "success",
            "name": None,
            "address": None,
            "email": None,
            "mobile": None
        }
        reader = PyPDF2.PdfReader(pdf_bytesio)
        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split("\n")
            for i, line in enumerate(lines):
                if re.search(r"\bRegistered\s+Architect\b", line, re.IGNORECASE):
                    name_candidate = None
                    for j in range(max(0, i-2), i):  
                        if "Thiru" in lines[j]:
                            name_candidate = lines[j].strip()
                            break
                    if not name_candidate and i > 0:
                        name_candidate = lines[i-1].strip()
                    data["name"] = name_candidate
                    block = lines[max(0, i-2):i+7]
                    block_text = "\n".join(block)
                    email_match = re.search(r"[\w\.-]+@[\w\.-]+", block_text)
                    if email_match:
                        data["email"] = email_match.group(0)
                    mobile_match = re.search(r"\b\d{10}\b", block_text)
                    if mobile_match:
                        data["mobile"] = mobile_match.group(0)
                    address_lines = []
                    for addr_line in lines[i+1:]: 
                        if data["email"] and data["email"] in addr_line:
                            continue
                        if data["mobile"] and data["mobile"] in addr_line:
                            continue
                        cleaned_line = re.sub(r"\b\d{10}\b", "", addr_line)
                        cleaned_line = re.sub(r"[\w\.-]+@[\w\.-]+", "", cleaned_line)
                        if cleaned_line.strip():
                            address_lines.append(cleaned_line.strip())
                            if re.search(r"\b\d{3}[-\s]?\d{3}\b", cleaned_line):
                                break
                    data["address"] = ", ".join(address_lines).strip(" ,")
                    return data
        return {"error": "No 'Registered Architect' section found.", "status": "failure"}
    except Exception as e:
        return {"error": str(e), "status": "failure"}