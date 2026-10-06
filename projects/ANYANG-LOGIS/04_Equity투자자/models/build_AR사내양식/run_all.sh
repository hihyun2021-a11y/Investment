set -e
cd "$(dirname "$0")"
python3 -I stage_a.py ../ar_in/form.xlsx stage_a.xlsx
python3 -I stage_b.py ../ar_in/model_v03.xlsm stage_a.xlsx v04_noval.xlsm
rm -rf lo_out && mkdir lo_out && cp v04_noval.xlsm lo_out/calc.xlsm
(cd lo_out && timeout 900 soffice -env:UserInstallation=file://$(pwd)/../lo_prof --headless --convert-to xlsx calc.xlsm >/dev/null 2>&1)
python3 -I collect.py
python3 -I stage_b.py ../ar_in/model_v03.xlsm stage_a.xlsx v04_final.xlsm values.json
python3 -I preview.py
soffice --headless --convert-to pdf preview.xlsx >/dev/null 2>&1
python3 -c "
import pymupdf
d=pymupdf.open('preview.pdf'); p=d[0]; r=p.rect
for i,(x0,x1,y0,y1) in enumerate([(0,0.36,0,0.62),(0.33,0.70,0,0.62),(0.62,1.0,0,0.62),(0.33,0.70,0.55,1.0)]):
    p.get_pixmap(dpi=200, clip=pymupdf.Rect(r.width*x0, r.height*y0, r.width*x1, r.height*y1)).save(f'z{i+1}.png')
"
