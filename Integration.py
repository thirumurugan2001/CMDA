from ZohoCRMAutomatedAuth import ZohoCRMAutomatedAuth
from helper import excel_to_json, assign_sales_person_to_areas, separate_and_store_temp, assgin_leads_to_lead_name, compare_and_update_excel

def lead_import(file_path):
    try: 
        update_success, update_data = compare_and_update_excel(new_file=file_path)        
        if not update_success:
            return {
                "message": "No new records to process. Excel file is up to date.",
                "statusCode": 200,
                "status": True,
                "analysis_data": update_data
            }        
        crm = ZohoCRMAutomatedAuth() 
        if crm.test_api_connection():
            matched_file_path, analysis_data = separate_and_store_temp(file_path, True)   
            if not matched_file_path:
                return {
                    "message": "Failed to separate records.",
                    "statusCode": 400,
                    "status": False,
                    "analysis_data": {}
                }
            area_result = assign_sales_person_to_areas(
                excel_file_path=matched_file_path,
                area_column_name="Area Name", 
                sales_person_column_name="Sales Person"
            )
            if isinstance(area_result, dict):
                analysis_data['unmatched_areas'] = area_result.get('unmatched_areas', [])
                matched_file_path = area_result.get('matched_file_path', matched_file_path)
            else:
                matched_file_path = area_result            
            analysis_data.update(update_data)            
            records = excel_to_json(matched_file_path)            
            if records: 
                cmda_success = crm.push_records_to_zoho(records)
                leads_success = assgin_leads_to_lead_name(area_result.get('matched_file_path', matched_file_path), crm)  
                print(f"CMDA Success: {cmda_success}, Leads Success: {leads_success}")         
                if cmda_success and leads_success:
                    return {
                        "message": "Records pushed to CMDA and Leads created successfully!",
                        "statusCode": 200,
                        "status": True,
                        "analysis_data": analysis_data
                    }
                else:
                    print("Some records failed to process. ********")
                    error_msg = []
                    if not cmda_success:
                        error_msg.append("Failed to push some CMDA records")
                    if not leads_success:
                        error_msg.append("Failed to create some Leads")
                    return {
                        "message": "/ ".join(error_msg),
                        "statusCode": 400,  
                        "status": False,
                        "analysis_data": analysis_data
                    }                    
            else:
                return {
                    "message": "No records found in Excel file",
                    "statusCode": 400,  
                    "status": False,
                    "analysis_data": analysis_data
                }
        else:
            return {
                "message": "API connection failed!",
                "statusCode": 400,  
                "status": False,
                "analysis_data": {}
            }        
    except Exception as e:
        print(f"Error in lead_import: {str(e)}")
        return {
            "message": str(e),
            "statusCode": 400,
            "status": False,
            "analysis_data": {}
        }