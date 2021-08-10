from selenium import webdriver
import time
from openpyxl import load_workbook
import subprocess
import webdriver_manager.chrome
import shutil

def main():
    '''My first python script, from a long time ago... Not pretty but still works eh
    scrapes hockey teams stats, and sends them to an excel, which "calculates odds" for upcoming matches'''
    liigat = {'NHL': 'jaakiekko/usa/nhl', 'KHL': 'jaakiekko/venaja/khl', 'LIIGA': 'liiga'}
    x = 0
    y = 8
    while True:
        liiga = input('Mikä liiga?\n')
        if liiga.upper() in liigat:
            valinta = liigat[liiga.upper()]
            break
        else:
            print('Ei kelvollinen liiga!')
    day = input('Minkä päivän pelit? (xx.yy.)\n')
    joukkueet = []
    parit = []
    arvot = []
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--disable-extensions')
    chrome_options.add_argument('--profile-directory=Default')
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--disable-plugins-discovery")
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(webdriver_manager.chrome.ChromeDriverManager().install(), options=chrome_options)
    driver.delete_all_cookies()
    driver.set_window_size(800, 800)
    driver.set_window_position(0, 0)
    driver.get(f'https://www.livetulokset.com/{valinta}/sarjataulukko/')
    time.sleep(3)
    nimi = driver.find_elements_by_xpath('//a[@target="_self"]')
    arvo = driver.find_elements_by_css_selector('span[class*="rowCell"]')
    for pala in nimi:
        if pala.text not in joukkueet:
            joukkueet.append(pala.text)
    time.sleep(2)
    for i in arvo:
        if len(arvot) < len(joukkueet)*8:
            if ':' in i.text:
                arvot.append(i.text.split(':')[0])
                arvot.append(i.text.split(':')[1])
            else:
                arvot.append(i.text)
    driver.find_element_by_xpath('//*[@id="onetrust-accept-btn-handler"]').click()
    driver.find_element_by_xpath('//a[@class="tabs__tab" and text()="Kunto"]').click()
    elem = driver.find_element_by_xpath('//a[@class="subTabs__tab " and text()="10"]')
    driver.execute_script("arguments[0].click();", elem)
    time.sleep(2)
    haku1 = 'div[class*="row_"]'
    pelit = driver.find_elements_by_css_selector(haku1)
    ls = []
    ls1 = {}
    for i in pelit:
        ls.append(':'.join(i.text.split('\n')))
    for a in ls:
        nimi = a.split(':')[1].upper()
        luvut = a.split(':')[3:9]
        if nimi not in ls1:
            ls1[nimi] = luvut

    driver.get(f'https://www.livetulokset.com/{valinta}/otteluohjelma/')
    time.sleep(2)
    haku = 'div[class*="event__match"]'
    ajat = driver.find_elements_by_css_selector(haku)

    for i in ajat:
        ls = i.text.split('\n')
        aika = ls[2].split(' ')[0]
        if aika == day:
            parit.append('{}:{}'.format(ls[0], ls[1]))
    time.sleep(3)
    driver.close()

    if len(parit) < 1:
        exit()
    tiimit = {}
    for i in joukkueet:
        tiimit[i.upper()] = arvot[x:y]

        x += 8
        y += 8

    gfka = 0
    gaka = 0

    for i in tiimit:
        gfka += int(tiimit[i][5])/int(tiimit[i][0])
        gaka += int(tiimit[i][6])/int(tiimit[i][0])

    gfka = gfka/len(joukkueet)
    gaka = gaka/len(joukkueet)

    kas = [0.0] * len(parit)
    kds = [0.0] * len(parit)
    vas = [0.0] * len(parit)
    vds = [0.0] * len(parit)

    kotiastr = [0.0] * len(parit)
    kotidstr = [0.0] * len(parit)
    vierasastr = [0.0] * len(parit)
    vierasdstr = [0.0] * len(parit)
    for i in range(len(parit)):
        koti = parit[i].split(':')[0].upper()
        vieras = parit[i].split(':')[1].upper()
        kotiastr[i] = (float(tiimit[koti][5])/float(tiimit[koti][0])) / gfka
        kotidstr[i] = (float(tiimit[koti][6])/float(tiimit[koti][0])) / gaka
        vierasastr[i] = (float(tiimit[vieras][5])/float(tiimit[vieras][0])) / gfka
        vierasdstr[i] = (float(tiimit[vieras][6])/float(tiimit[vieras][0])) / gaka
        kas[i] = (float(ls1[koti][4]) / 10.0) / gfka
        kds[i] = (float(ls1[koti][5]) / 10.0) / gaka
        vas[i] = (float(ls1[vieras][4]) / 10.0) / gfka
        vds[i] = (float(ls1[vieras][5]) / 10.0) / gaka

    filu = 'piv3.xlsx'
    wb = load_workbook(filu)
    ws = wb.worksheets[0]
    if len(wb.worksheets) < len(parit):
        for i in range(len(parit)):
            if len(wb.worksheets) < len(parit):
                wb.copy_worksheet(ws)
    elif len(wb.worksheets) > len(parit):
        on = wb.worksheets
        for i in range(len(wb.worksheets)):
            if i + 1 > len(parit):
                wb.remove(on[i])
    for i in range(len(parit)):
        koti = parit[i].split(':')[0].upper()
        vieras = parit[i].split(':')[1].upper()
        sht = wb.worksheets[i]
        sht['A2'] = kotiastr[i]
        sht['A3'] = vierasastr[i]
        sht['B2'] = kotidstr[i]
        sht['B3'] = vierasdstr[i]
        sht['A6'] = kotiastr[i] * vierasdstr[i] * gaka
        sht['A7'] = kotidstr[i] * vierasastr[i] * gaka
        sht['C9'] = '{} (koti)'.format(koti)
        sht['D8'] = '{} (vieras)'.format(vieras)
        sht['K2'] = kas[i]
        sht['K3'] = vas[i]
        sht['L2'] = kds[i]
        sht['L3'] = vds[i]
        sht['K6'] = kas[i] * vds[i] * gaka
        sht['K7'] = kds[i] * vas[i] * gaka
    wb.save('piv3.xlsx')
    shutil.copy('piv3.xlsx', '{}{}.xlsx'.format(day, liiga))
    subprocess.Popen('piv3.xlsx', shell=True)



if __name__ == "__main__":
    main()
