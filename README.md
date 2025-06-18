PySimpleGUI je v poslední verzi placený, proto byla použita starší verze
kvůli starší verzi je vyžadován python 11
pip install co je na odkazu nejspíše nebude fungovat, proto je potřeba soubor stáhnout

PySimpleGUI:
https://pypi.org/project/PySimpleGUI-4-foss/#files

ideální je .whl soubor dát do složky ve které je projekt a píše se do konzole


INSTALACE:
ve složce připravené pro projekt:
git clone https://github.com/hajloun/Excel-to-Excel.git
py -3.11 -m venv NAZEV_VIRTUALNIHO_PROSTREDI
.\venv\Scripts\activate
cd .\Excel-to-Excel\
pip install -r requirements.txt
pip install PySimpleGUI_4_foss-4.60.4.1-py3-none-any.whl

spuštění programu:
python main.py

při každém zapnutí VSC je potřeba aktivovat virtuální prostředí
.\venv\Scripts\activate
