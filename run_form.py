import subprocess
import os
import sys

if __name__ == "__main__":
    base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(base_dir, "form_app.py")
    print("Iniciando Gerador de Relatorios BOMPARC...")
    
    # Chama o python local e utiliza a flag -m para chamar o streamlit (evitando erros de PATH do Windows)
    subprocess.run(f'python -m streamlit run "{script_path}"', shell=True)
