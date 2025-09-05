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

# ------------------- Upload API -------------------
@api_view(['POST'])
def upload_files(request):
    job_id = str(uuid.uuid4())
    folder = _job_dir(job_id)
    folder.mkdir(parents=True, exist_ok=True)

    files = request.FILES.getlist('files')
    if not files:
        return HttpResponseBadRequest("No files uploaded")

    for f in files:
        filename = f.name
        path = folder / filename
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, 'wb+') as dest:
            for chunk in f.chunks():
                dest.write(chunk)

    JOBS[job_id] = {"progress": 0, "done": False, "result": None, "error": None}
    return Response({"jobId": job_id})

# ------------------- Worker function -------------------
def _convert_worker(job_id: str):
    try:
        JOBS[job_id]["progress"] = 10
        folder = _job_dir(job_id)

        all_data = []
        for file in os.listdir(folder):
            if not file.endswith(".docx") or file.startswith("~$"):
                continue
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
            for i, chunk in enumerate(chunks, start=1):
                row_data[f"Description_Part{i}"] = chunk
            all_data.append(row_data)

        # save to Excel
        df = pd.DataFrame(all_data)
        xlsx_path = folder / "result.xlsx"
        csv_path = folder / "result.csv"
        df.to_excel(xlsx_path, index=False)
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")

        JOBS[job_id]["result"] = {"xlsx": str(xlsx_path), "csv": str(csv_path)}
        JOBS[job_id]["progress"] = 100
        JOBS[job_id]["done"] = True

    except Exception as e:
        JOBS[job_id]["error"] = str(e)
        JOBS[job_id]["done"] = True

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

    if fmt == "csv":
        return FileResponse(open(path, "rb"), as_attachment=True, filename="result.csv", content_type="text/csv")
    else:
        return FileResponse(open(path, "rb"), as_attachment=True, filename="result.xlsx",
                            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
