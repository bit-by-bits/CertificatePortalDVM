from django.shortcuts import render, redirect
from django.core.files.storage import default_storage
from django.http import HttpResponse

from .cert import generate
import os
from openpyxl import load_workbook
from CertificatePortalDVM.settings import MEDIA_ROOT
import shutil


def index(request):
    return render(request, 'portal/index.html')


def cert_portal(request):
    if request.method == "POST":
        file = request.FILES['excel_file']
        file_name = default_storage.save(file.name, file)
        loc = (os.path.join(MEDIA_ROOT, f'./{file_name}'))
        wb_read = load_workbook(filename=loc)
        sheet = wb_read.active

        generate(sheet)

        # TODO 126 Dhruv Garg gives an error possibly because of '\t\ char messing with the file path

        zipped_certificates = shutil.make_archive(f'{MEDIA_ROOT}/Certs', 'zip', f'{MEDIA_ROOT}/./Certificates')
        with open(f'{MEDIA_ROOT}/./Certs.zip', 'rb') as f:
            zipped_certificates = f.read()
        response = HttpResponse(zipped_certificates, content_type='application/force-download')
        response['Content-Disposition'] = 'attachment; filename="Certificates.zip"'

        os.remove(os.path.join(MEDIA_ROOT, './Certs.zip'))
        shutil.rmtree(os.path.join(MEDIA_ROOT, './Certificates'))
        os.remove(os.path.join(MEDIA_ROOT, file_name))

        return response

    return redirect('CertPortal:portal')
