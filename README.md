nombre_del_archivo = 'SOLICITUDES 11 DE JUNIO.xlsx'  #@param {type:"string"}
from google.colab import drive
drive.mount('/content/drive')
!pip install --upgrade openpyxl
!curl -sSL "https://raw.githubusercontent.com/frtobarn/solicitudes-logic/refs/heads/main/logic.py" -o lg.py
!python3 lg.py "{nombre_del_archivo}"
