import json,re
from pathlib import Path
p=Path('work/turnover-monitor-preview.html')
s=p.read_text(encoding='utf-8')
map={}
def add(codes, path):
    for code in codes.split(): map[code]=path
foam='https://storage.yandexcloud.net/kronbuild-assets/%D0%9F%D0%B5%D0%BD%D1%8B%20%D0%BC%D0%BE%D0%BD%D1%82%D0%B0%D0%B6%D0%BD%D1%8B%D0%B5%20KronBuild/'
add('FH200',foam+'KronBuild_%D0%91%D1%8B%D1%82%D0%BE%D0%B2%D0%B0%D1%8F%2020%2B.png')
add('FH400',foam+'KronBuild_%D0%91%D1%8B%D1%82%D0%BE%D0%B2%D0%B0%D1%8F%2040%2B.png')
add('FHO65',foam+'KronBuild_%D0%91%D1%8B%D1%82%D0%BE%D0%B2%D0%B0%D1%8F_1%20(One)%2066.png')
add('FAS650',foam+'KronBuild_%D0%92%D1%81%D0%B5%D1%81%D0%B5%D0%B7%D0%BE%D0%BD%D0%BD%D0%B0%D1%8F%2065.png')
add('FAO65',foam+'KronBuild_1%20(One)%2065.png')
add('FASGUN65',foam+'KronBuild_%D0%A3%D0%BD%D0%B8%D0%B2%D0%B5%D1%80%D1%81%D0%B0%D0%BB65.png')
add('FSS65 FWS65 FASST',foam+'KronBuild_%D0%A1%D1%82%D0%B0%D0%BD%D0%B4%D0%B0%D1%80%D1%82%2065_%D0%B2%D1%81%D0%B5%D1%81%D0%B5%D0%B7%D0%BE%D0%BD%D0%BD%D0%B0%D1%8F.png')
add('FSU65 FWU65 FAS65U',foam+'KronBuild_%D0%A3%D0%BB%D1%8C%D1%82%D1%80%D0%B065_%D0%B2%D1%81%D0%B5%D1%81%D0%B5%D0%B7%D0%BE%D0%BD%D0%BD%D0%B0%D1%8F.png')
add('FPFB1',foam+'KronBuild_%D0%9E%D0%B3%D0%BD%D0%B5%D1%81%D1%82%D0%BE%D0%B9%D0%BA%D0%B0%D1%8F%2065.png')
add('GF800',foam+'KronBuild_%D0%9A%D0%BB%D0%B5%D0%B9_%D0%BF%D0%B5%D0%BD%D0%B0.png')
add('FC850',foam+'KronBuild_%D0%A6%D0%B5%D0%BC%D0%B5%D0%BD%D1%82_%D0%BF%D0%B5%D0%BD%D0%B0.png')
add('PUI90',foam+'KronBuild_%D0%A3%D1%82%D0%B5%D0%BF%D0%BB%D0%B8%D1%82%D0%B5%D0%BB%D1%8C.png')
add('CF351',foam+'KronBuild_%D0%9E%D1%87%D0%B8%D1%81%D1%82%D0%B8%D1%82%D0%B5%D0%BB%D1%8C.png')
add('FSPM80 FWPM80',foam+'KronBuild_%D0%AD%D0%BA%D1%81%D0%BF%D0%B5%D1%80%D1%82_%D0%9C%D0%B0%D1%81%D1%82%D0%B5%D1%8080_%D0%B7%D0%B8%D0%BC%D0%B0.png')
add('FSP70 FWP70 FAS70',foam+'KronBuild_%D0%AD%D0%BA%D1%81%D0%BF%D0%B5%D1%80%D1%82_%D0%9F%D1%80%D0%BE70_%D0%B2%D1%81%D0%B5%D1%81%D0%B5%D0%B7%D0%BE%D0%BD%D0%BD%D0%B0%D1%8F.png')
add('SO65',foam+'KronBuild_%D0%A1%D1%82%D0%B0%D1%80%D1%82_One65.png')
add('SS65',foam+'KronBuild_%D0%A1%D1%82%D0%B0%D1%80%D1%82_C%D1%82%D0%B0%D0%BD%D0%B4%D0%B0%D1%80%D1%8265.png')
add('SP70',foam+'KronBuild_%D0%A1%D1%82%D0%B0%D1%80%D1%82_%D0%9F%D1%80%D0%BE70.png')
add('SF65',foam+'KronBuild_%D0%A1%D1%82%D0%B0%D1%80%D1%82_%D0%9E%D0%B3%D0%BD%D0%B5%D1%81%D1%82%D0%BE%D0%B9%D0%BA%D0%B0%D1%8F%2065.png')
add('SKP',foam+'KronBuild_%D0%A1%D1%82%D0%B0%D1%80%D1%82_%D0%9A%D0%BB%D0%B5%D0%B9_%D0%BF%D0%B5%D0%BD%D0%B0.png')
add('SNU',foam+'KronBuild_%D0%A1%D1%82%D0%B0%D1%80%D1%82_%D0%9D%D0%B0%D0%BF%D1%8B%D0%BB%D1%8F%D0%B5%D0%BC%D1%8B%D0%B9_%D1%83%D1%82%D0%B5%D0%BF%D0%BB%D0%B8%D1%82%D0%B5%D0%BB%D1%8C.png')
sg='https://storage.yandexcloud.net/kronbuild-assets/%D0%93%D0%B5%D1%80%D0%BC%D0%B5%D1%82%D0%B8%D0%BA%D0%B8%20%D0%B8%20%D0%BA%D0%BB%D0%B5%D0%B9%20KronBuild/'
for codes,fn in {
'SUT80 SUW80 SUW28 SUT28 SUB28':'%D0%A3%D0%BD%D0%B8%D0%B2%D0%B5%D1%80%D0%B5%D1%81%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9.png','SST80 SSW80 SSW28 SST28':'%D0%A1%D0%B0%D0%BD%D0%B8%D1%82%D0%B0%D1%80%D0%BD%D1%8B%D0%B9.png','SNT60 SNW28 SNT28':'%D0%9D%D0%B5%D0%B9%D1%82%D1%80%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9.png','SHR80 SHR30':'%D0%9E%D0%B3%D0%BD%D0%B5%D1%81%D1%82%D0%BE%D0%B9%D0%BA%D0%B8%D0%B9.png','SAT30 SAB30':'%D0%90%D0%BA%D0%B2%D0%B0%D1%83%D0%B8%D1%83%D0%BC%D0%BD%D1%8B%D0%B9.png','WS650':'WS650.png','ASA28':'%D0%90%D0%BA%D1%80%D0%B8%D0%BB%D0%BE%D0%B2%D1%8B%D0%B9.png','PSG78 PSB78 PSW78':'PU40.png','MSC36':'%D0%9A%D1%80%D0%B8%D1%81%D1%82%D0%B0%D0%BB.png','LNB39':'%D0%96%D0%B8%D0%B4%D0%BA%D0%B8%D0%B5%20%D0%B3%D0%B2%D0%BE%D0%B7%D0%B4%D0%B8_%D0%B1%D0%B5%D0%B6.png','LNW40':'%D0%96%D0%B8%D0%B4%D0%BA%D0%B8%D0%B5%20%D0%B3%D0%B2%D0%BE%D0%B7%D0%B4%D0%B8_%D0%B1%D0%B5%D0%BB.png','EXP06 EXP25':'%D0%AD%D0%BF%D0%BE%D0%BA%D1%81%D0%B8%D0%B4%D0%BD%D1%8B%D0%B9%20%D0%BA%D0%BB%D0%B5%D0%B9%20%D0%BC%D0%B8%D0%BD%D0%B8.png','CBR20 GSX03':'%D0%9A%D0%BB%D0%B5%D0%B9.png','SMPSW78 SMPSG78 SMPSP78':'%D0%A1%D1%82%D0%B0%D1%80%D1%82_PU40.png','SMSUW28 SMSUT28':'%D0%A1%D1%82%D0%B0%D1%80%D1%82_%D1%83%D0%BD%D0%B8%D0%B2%D0%B5%D1%80%D1%81%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9.png','SMSSW28 SMSST28':'%D0%A1%D1%82%D0%B0%D1%80%D1%82_%D1%81%D0%B0%D0%BD%D0%B8%D1%82%D0%B0%D1%80%D0%BD%D1%8B%D0%B9.png'}.items(): add(codes,sg+fn)
m=re.search(r'const data=(\[.*?\]);',s)
data=json.loads(m.group(1)); changed=0
for x in data:
    a=str(x.get('article','')).strip()
    if a in map:
        x['image']=map[a]; x['imageSource']='google_sheet'; changed+=1
s=s[:m.start(1)]+json.dumps(data,ensure_ascii=False,separators=(',',':'))+s[m.end(1):]
s=s.replace('Фото взяты из каталога KRONbuild','Фото пены, герметиков и клеёв взяты из Google-таблицы KRONbuild')
p.write_text(s,encoding='utf-8')
print('updated',changed,'cards; mapped',len(map))
