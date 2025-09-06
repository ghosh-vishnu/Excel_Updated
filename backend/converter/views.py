from django.shortcuts import render

# Create your views here.
import os, uuid, threading
from pathlib import Path
from django.conf import settings
from django.http import FileResponse, Http404, HttpResponseBadRequest
from rest_framework.decorators import api_view
from rest_framework.response import Response
from datetime import date
import pandas as pd


from converter.utils import extractor

# simple in-memory job tracker
JOBS = {}

def _job_dir(job_id: str) -> Path:
    return Path(settings.MEDIA_ROOT) / job_id

def _cleanup_uploaded_files(folder: Path):
    """Clean up uploaded Word files after successful conversion, keeping only the output files."""
    try:
        for file_path in folder.iterdir():
            if file_path.is_file():
                # Keep only the output Excel and CSV files
                if not (file_path.name.endswith('.xlsx') or file_path.name.endswith('.csv')):
                    file_path.unlink()  # Delete the file
                    print(f"Cleaned up: {file_path.name}")
    except Exception as e:
        print(f"Error during cleanup: {e}")

def _cleanup_old_jobs():
    """Clean up old job folders that are no longer needed."""
    try:
        import time
        current_time = time.time()
        media_root = Path(settings.MEDIA_ROOT)
        
        for job_folder in media_root.iterdir():
            if job_folder.is_dir() and job_folder.name not in JOBS:
                # Check if folder is older than 1 hour (3600 seconds)
                folder_age = current_time - job_folder.stat().st_mtime
                if folder_age > 3600:  # 1 hour
                    import shutil
                    shutil.rmtree(job_folder)
                    print(f"Cleaned up old job folder: {job_folder.name}")
    except Exception as e:
        print(f"Error during old job cleanup: {e}")

# ------------------- Upload API -------------------
@api_view(['POST'])
def upload_files(request):
    # Clean up old job folders before starting new upload
    _cleanup_old_jobs()
    
    job_id = str(uuid.uuid4())
    folder = _job_dir(job_id)
    folder.mkdir(parents=True, exist_ok=True)

    files = request.FILES.getlist('files')
    if not files:
        return HttpResponseBadRequest("No files uploaded")

    # Extract folder name from the first file's webkitRelativePath
    folder_name = "Word_Files"  # Default name
    if files and hasattr(files[0], 'name'):
        # Try to get folder name from webkitRelativePath if available
        webkit_path = getattr(files[0], 'webkitRelativePath', None)
        if webkit_path and '/' in webkit_path:
            folder_name = webkit_path.split('/')[0]
        elif webkit_path and '\\' in webkit_path:
            folder_name = webkit_path.split('\\')[0]
        else:
            # If no webkitRelativePath, use a sanitized version of the first file's directory
            folder_name = "Word_Files"

    # Sanitize folder name for filename
    import re
    folder_name = re.sub(r'[^\w\s-]', '', folder_name).strip()
    folder_name = re.sub(r'[-\s]+', '_', folder_name)
    if not folder_name:
        folder_name = "Word_Files"

    for f in files:
        filename = f.name
        path = folder / filename
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, 'wb+') as dest:
            for chunk in f.chunks():
                dest.write(chunk)

    JOBS[job_id] = {"progress": 0, "done": False, "result": None, "error": None, "folder_name": folder_name}
    return Response({"jobId": job_id})

# ------------------- Worker function -------------------
def _convert_worker(job_id: str):
    try:
        JOBS[job_id]["progress"] = 5
        folder = _job_dir(job_id)

        # Get list of files to process
        files_to_process = [f for f in os.listdir(folder) if f.endswith(".docx") and not f.startswith("~$")]
        total_files = len(files_to_process)
        
        if total_files == 0:
            JOBS[job_id]["progress"] = 100
            JOBS[job_id]["done"] = True
            return

        all_data = []
        for i, file in enumerate(files_to_process):
            path = folder / file
            print(f"Processing {file}...")

            # run all extractors
            title = extractor.extract_title(str(path))
            description = extractor.extract_description(str(path))
            toc = extractor.extract_toc(str(path))
            methodology = extractor.extract_methodology_from_faqschema(str(path))
            seo_title = extractor.extract_seo_title(str(path))
            breadcrumb_text = extractor.extract_breadcrumb_text(str(path))
            skucode = extractor.extract_sku_code(str(path))
            urlrp = extractor.extract_sku_url(str(path))
            breadcrumb_schema = extractor.extract_breadcrumb_schema(str(path))
            meta = extractor.extract_meta_description(str(path))
            schema2 = extractor.extract_faq_schema(str(path))
            report = extractor.extract_report_coverage_table_with_style(str(path))
            merge = extractor.merge_description_and_coverage(str(path))
            chunks = extractor.split_into_excel_cells(merge)

            row_data = {
                "File": file,
                "Title": title,
                "Description": description,
                "TOC": toc,
                "Segmentation": "<p>.</p>",
                "Methodology": methodology,
                "Publish_Date": date.today().strftime("%B %Y"),
                "Currency": "USD",
                "Single Price": 4485,
                "Corporate Price": 6449,
                "skucode": skucode,
                "Total Page": "",
                "Date": date.today().strftime("%Y-%m-%d"),
                "urlNp": urlrp,
                "Meta Description": meta,
                "Meta Keys": "",
                "Base Year": "2024",
                "history": "2019-2023",
                "Enterprise Price": 8339,
                "SEOTITLE": seo_title,
                "BreadCrumb Text": breadcrumb_text,
                "Schema 1": breadcrumb_schema,
                "Schema 2": schema2,
                "Report": report,
                "Description_Merged": merge
            }
            for j, chunk in enumerate(chunks, start=1):
                row_data[f"Description_Part{j}"] = chunk
            all_data.append(row_data)
            
            # Update progress after each file (80% for file processing, 20% for final steps)
            file_progress = 5 + int((i + 1) / total_files * 80)
            JOBS[job_id]["progress"] = file_progress

        # save to Excel
        df = pd.DataFrame(all_data)
        folder_name = JOBS[job_id].get("folder_name", "Word_Files")
        xlsx_filename = f"{folder_name}.xlsx"
        csv_filename = f"{folder_name}.csv"
        xlsx_path = folder / xlsx_filename
        csv_path = folder / csv_filename
        df.to_excel(xlsx_path, index=False)
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")

        JOBS[job_id]["result"] = {"xlsx": str(xlsx_path), "csv": str(csv_path)}
        JOBS[job_id]["progress"] = 100
        JOBS[job_id]["done"] = True
        
        # Clean up uploaded files after successful conversion
        _cleanup_uploaded_files(folder)

    except Exception as e:
        JOBS[job_id]["error"] = str(e)
        JOBS[job_id]["done"] = True
        
        # Clean up uploaded files even if conversion failed
        _cleanup_uploaded_files(folder)

# ------------------- Start Conversion -------------------
@api_view(['POST'])
def start_convert(request):
    job_id = request.GET.get("jobId")
    if not job_id or job_id not in JOBS:
        return HttpResponseBadRequest("Invalid jobId")

    t = threading.Thread(target=_convert_worker, args=(job_id,), daemon=True)
    t.start()
    return Response({"started": True})

# ------------------- Progress -------------------
@api_view(['GET'])
def progress(request):
    job_id = request.GET.get("jobId")
    if not job_id or job_id not in JOBS:
        raise Http404("job not found")
    return Response(JOBS[job_id])

# ------------------- Result download -------------------
@api_view(['GET'])
def result_file(request):
    job_id = request.GET.get("jobId")
    if not job_id or job_id not in JOBS:
        raise Http404("job not found")

    fmt = request.GET.get("format", "xlsx").lower()
    path = JOBS[job_id].get("result", {}).get(fmt)
    if not path or not os.path.exists(path):
        raise Http404("result not ready")

    # Get the folder name for the download filename
    folder_name = JOBS[job_id].get("folder_name", "Word_Files")
    
    if fmt == "csv":
        filename = f"{folder_name}.csv"
        return FileResponse(open(path, "rb"), as_attachment=True, filename=filename, content_type="text/csv")
    else:
        filename = f"{folder_name}.xlsx"
        return FileResponse(open(path, "rb"), as_attachment=True, filename=filename,
                            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
