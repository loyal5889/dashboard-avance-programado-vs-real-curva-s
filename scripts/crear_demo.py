"""Crea fuentes ficticias compatibles con las consultas del modelo; no usa Project."""
from pathlib import Path
from datetime import date, datetime, timedelta
import ast, csv, argparse
from openpyxl import Workbook


def crear(destino):
    root=Path(__file__).resolve().parents[1]
    tree=ast.parse((root/'src/extractor_project.py').read_text(encoding='utf-8'))
    columnas=next(ast.literal_eval(n.value) for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='COLUMNS' for t in n.targets))
    target=destino/'output/cronograma_darca2.csv'
    if target.exists() or list((destino/'avance_real').glob('*.xlsx')):
        raise SystemExit('Ya hay fuentes en el destino. Elija una carpeta vacía para no sobrescribirlas.')
    target.parent.mkdir(parents=True,exist_ok=True);(destino/'avance_real').mkdir(parents=True,exist_ok=True)
    inicio=datetime(2026,1,5,8)
    records=[]
    for i in range(1,7):
        start=inicio+timedelta(days=(i-1)*3);finish=start+timedelta(days=23)
        r={c:'' for c in columnas}
        r.update(ProjectName='Proyecto ficticio',ExtractionDate='2026-02-14 12:00:00',StatusDate='2026-02-14 12:00:00',HoursPerDay=8,HoursPerWeek=48,MinutesPerDay=480,MinutesPerWeek=2880,ID=i,UniqueID=i,WBS=f'1.1.{i}',OutlineLevel=3,TaskName=f'Actividad demo {i}',Start=start.strftime('%Y-%m-%d %H:%M:%S'),Finish=finish.strftime('%Y-%m-%d %H:%M:%S'),BaselineStart=start.strftime('%Y-%m-%d %H:%M:%S'),BaselineFinish=finish.strftime('%Y-%m-%d %H:%M:%S'),DurationMinutes=9600,DurationDays=20,BaselineDurationMinutes=9600,BaselineDurationDays=20,PercentComplete=70,PhysicalPercentComplete=70,ActualStart=start.strftime('%Y-%m-%d %H:%M:%S'),ActualDurationMinutes=6720,ActualDurationDays=14,RemainingDurationMinutes=2880,RemainingDurationDays=6,Critical='false',Summary='false',Milestone='false',Active='true',BaselineAvailable='true',ConstraintType=0,ConstraintTypeName='As Soon As Possible',Cost=1000000,BaselineCost=1000000)
        for c in columnas:
            if r[c]=='' and (c.endswith('Minutes') or c.endswith('Days')):r[c]=0
        records.append(r)
    with target.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=columnas,delimiter=';');w.writeheader();w.writerows(records)
    for week,qty in enumerate((10,20,25,15),1):
        wb=Workbook();ws=wb.active;ws.title='DATOS';ws.append(['EJEMPLO SINTÉTICO']);ws.append([]);ws.append([])
        ws.append(['Fecha Corte','ID','WBS','Actividad','Tramo','Unidad','Cantidad Total','Cantidad Ejecutada Semana','Observación'])
        for r in records:ws.append([date(2026,1,10)+timedelta(days=(week-1)*7),r['ID'],r['WBS'],r['TaskName'],'TRAMO 1','m²',100,qty,'Datos ficticios'])
        wb.save(destino/f'avance_real/semana_{week:02}.xlsx');wb.close()
    return target


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--destino',type=Path,default=Path(__file__).resolve().parents[1]/'data/demo')
    print(crear(parser.parse_args().destino))
