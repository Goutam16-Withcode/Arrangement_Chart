# seating/views.py
import subprocess
from django.shortcuts import render
from django.http import JsonResponse
from django.shortcuts import render

def seating_view(request):
    return render(request, 'seating/index.html') 

def run_script(request):
    if request.method == "POST":
        try:
           
            result = subprocess.run(
                ['python', 'aarna.py'],  
                capture_output=True,
                text=True,
                check=True
            )
            return JsonResponse({'output': result.stdout, 'error': None})
        except subprocess.CalledProcessError as e:
            return JsonResponse({'output': None, 'error': e.stderr})