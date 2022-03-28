import time

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
        unique_time_stamp = str(float(time.time())).replace('.', '')
        file = request.FILES['excel_file']
        file_name = default_storage.save(file.name, file)
        loc = (os.path.join(MEDIA_ROOT, f'./{file_name}'))
        wb_read = load_workbook(filename=loc)
        sheet = wb_read.active

        generate(sheet, unique_time_stamp)

        zipped_certificates = shutil.make_archive(f'{MEDIA_ROOT}/Certs{unique_time_stamp}', 'zip', f'{MEDIA_ROOT}/./Certificates{unique_time_stamp}')
        # print(f"zipped! {zipped_certificates}")
        with open(f'{MEDIA_ROOT}/./Certs{unique_time_stamp}.zip', 'rb') as f:
            zipped_certificates = f.read()
        response = HttpResponse(zipped_certificates, content_type='application/force-download')
        response['Content-Disposition'] = f'attachment; filename="Certificates{unique_time_stamp}.zip"'

        try:
            os.remove(os.path.join(MEDIA_ROOT, f'./Certs{unique_time_stamp}.zip'))
            # print(f"removed Certs{unique_time_stamp}.zip")
            shutil.rmtree(os.path.join(MEDIA_ROOT, f'./Certificates{unique_time_stamp}'))
            # print(f"removed Certificates{unique_time_stamp}")
            os.remove(os.path.join(MEDIA_ROOT, file_name))
            # print(f"removed winners.xlsx")
        except PermissionError:
            pass
            # using this to not stop the program execution and handling WinError32 (Permission Error)
            # this error is probably fixed since timestamp will give it a unique name but even then if there is an error this will prevent the server from stoppping.
        return response

    return redirect('CertPortal:portal')
