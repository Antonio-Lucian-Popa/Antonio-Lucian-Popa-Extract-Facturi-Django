import os
import uuid
from django.conf import settings
from django.http import FileResponse
from django.shortcuts import render
from django.views.decorators.csrf import csrf_exempt
from extrage_facturi import process_pdfs, finalize_excel

UPLOAD_DIR = os.path.join(settings.BASE_DIR, "media", "temp_uploads")
RESULT_DIR = os.path.join(settings.BASE_DIR, "media", "temp_results")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)

@csrf_exempt
def upload_view(request):
    if request.method == 'POST':
        files = request.FILES.getlist('pdf_files')
        if not files:
            return render(request, 'procesare/upload.html', {'error': 'Nu ai trimis fișiere'})

        # Șterge fișiere vechi
        for f in os.listdir(UPLOAD_DIR):
            os.remove(os.path.join(UPLOAD_DIR, f))

        for f in files:
            path = os.path.join(UPLOAD_DIR, f.name)
            with open(path, 'wb+') as dest:
                for chunk in f.chunks():
                    dest.write(chunk)

        # Generează fișier excel
        output_filename = f"rezultate_facturi_{uuid.uuid4().hex[:6]}.xlsx"
        output_path = os.path.join(RESULT_DIR, output_filename)

        rows = process_pdfs(UPLOAD_DIR, output_path)
        finalize_excel(rows, output_path)

        return FileResponse(open(output_path, 'rb'), as_attachment=True, filename=output_filename)

    return render(request, 'procesare/upload.html')
