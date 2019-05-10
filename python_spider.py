import requests
from bs4 import BeautifulSoup as bs
from openpyxl import workbook



headers = {
"user-agent":"Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Mobile Safari/537.36"
}

def getHtml(ID,SD,END):
    id = ID
    url1 = "http://sdhq.hust.edu.cn/ICBS/PurchaseWebService.asmx/getMeterDayValue?"

    parse = {
    "AmMeter_ID":ID,
    "startDate":SD,
    "endDate":END
    }
    re = requests.get(url = url1,params = parse,headers = headers)
    webpage = re.content
    getData(webpage)

def getData(webpage):
    global ws
    data = []  #  存储电量数据
    time = []  #  存储时间
    soup = bs(webpage,"html.parser")

    for d in soup.find_all('dayvalue'):
        da = d.get_text()
        data.append(da)

    for t in soup.find_all('curdaytime'):
        ti = t.get_text()
        time.append(ti)
    for i in range(30):
        ws.append([data[i],time[i]])

SD = "2019/4/1"
END = "2019/4/30"


    # ID6# 阳面
IDS6 = [
    '1001.02470007.1',
    '1001.02470011.1',
    '1001.02470027.1',
    '1001.02470033.1',
    '1001.02470132.1',
    '1001.02470134.1',
    '1001.02470085.1',
    '1001.02470087.1',
    '1001.02470145.1',
    '1001.02470149.1',
    '1001.02470168.1',
    '1001.02470171.1'

    ]

roomS6 = [
    '101',
    '105',
    '119',
    '125',
    '401',
    '403',
    '423',
    '425',
    '601',
    '605',
    '622',
    '625'
    ]
    # 阴面
IDN6 = [
    '1001.02470008.1',
    '1001.02470012.1',
    '1001.02470026.1',
    '1001.02470032.1',
    '1001.02470133.1',
    '1001.02470135.1',
    '1001.02470084.1',
    '1001.02470088.1',
    '1001.02470146.1',
    '1001.02470148.1',
    '1001.02470169.1',
    '1001.02470172.1'

    ]

roomN6 = [
    '102',
    '106',
    '118',
    '124',
    '402',
    '404',
    '422',
    '426',
    '602',
    '604',
    '623',
    '626'
    ]

    # ID7# 阳面
IDS7 = [
    '1001.02480145.1',
    '1001.02480149.1',
    '1001.02480008.1',
    '1001.02480022.1',
    '1001.02480218.1',
    '1001.02480220.1',
    '1001.02480082.1',
    '1001.02480094.1',
    '1001.02480265.1',
    '1001.02480273.1',
    '1001.02480137.1',
    '1001.02480141.1'
    ]

roomS7 = [
    '101',
    '105',
    '129',
    '141',
    '401',
    '403',
    '431',
    '空441',
    '601',
    '609',
    '637',
    '641'
    ]
    # 阴面
IDN7 = [
    '1001.02480146.1',
    '1001.02480150.1',
    '1001.02480007.1',
    '1001.02480021.1',
    '1001.02480219.1',
    '1001.02480221.1',
    '1001.02480081.1',
    '1001.02480095.1',
    '1001.02480266.1',
    '1001.02480272.1',
    '1001.02480138.1',
    '1001.02480142.1'

    ]

roomN7 = [
    '102',
    '106',
    '128',
    '140',
    '402',
    '404',
    '430',
    '442',
    '602',
    '608',
    '638',
    '642'
    ]

    # ID8# 阳面
IDS8 = [
    '1001.02490001.1',
    '1001.02480150.1',
    '1001.02480007.1',
    '1001.02480021.1',
    '1001.02480219.1',
    '1001.02480221.1',
    '1001.02480081.1',
    '1001.02480095.1',
    '1001.02480266.1',
    '1001.02480272.1',
    '1001.02480138.1',
    '1001.02480142.1'

    ]

roomS8 = [
    '空101',
    '117',
    '1？？',
    '119',
    '401',
    '405',
    '415',
    '419',
    '601',
    '607',
    '617',
    '619'
    ]

    # 阴面
IDN8 = [
    '1001.02490002.1',
    '1001.02490015.1',
    '1001.02480008.1',
    '1001.02490017.1',
    '1001.02490062.1',
    '1001.02490066.1',
    '1001.02490078.1',
    '1001.02490082.1',
    '1001.02470146.1',
    '1001.02490116.1',
    '1001.02490124.1',
    '1001.02490144.1'

    ]

roomN8 = [
    '空102',
    '116',
    '1？？',
    '118',
    '402',
    '406',
    '416',
    '420',
    '602',
    '608',
    '616',
    '620'
    ]

    # ID9# 阳面_一楼无参考价值
IDS9 = [
    '1001.02500080.1',
    '1001.02500090.1',
    '1001.02500100.1',
    '1001.02500104.1',
    '1001.02500135.1',
    '1001.02500137.1',
    '1001.02500156.1',
    '1001.02500160.1'

    ]

roomS9 = [
    '401',
    '411',
    '421',
    '425',
    '601',
    '603',
    '621',
    '625'
    ]
    # 阴面_一楼无参考价值
IDN9 = [
    '1001.02500081.1',
    '1001.02500091.1',
    '1001.02500101.1',
    '1001.02500105.1',
    '1001.02500136.1',
    '1001.02500138.1',
    '1001.02500157.1',
    '1001.02500161.1'

    ]

roomN9 = [
    '402',
    '412',
    '422',
    '426',
    '602',
    '604',
    '622',
    '626'

    ]

    # ID10# 阳面
IDS10 = [
    '1001.02510001.1',
    '1001.02510011.1',
    '1001.02510019.1',
    '1001.02510025.1',
    '1001.02510080.1',
    '1001.02510090.1',
    '1001.02510098.1',
    '1001.02510104.1',
    '1001.02510133.1',
    '1001.02510143.1',
    '1001.02510155.1',
    '1001.02510146.1'

    ]
roomS10 = [
    '101',
    '111',
    '119',
    '125',
    '401',
    '411',
    '419',
    '425',
    '601',
    '611',
    '619',
    '625'

    ]
    # 阴面
IDN10 = [
    '1001.02510002.1',
    '1001.02510010.1',
    '1001.02510020.1',
    '1001.02510024.1',
    '1001.02510081.1',
    '1001.02510089.1',
    '1001.02510099.1',
    '1001.02510105.1',
    '1001.02510134.1',
    '1001.02510142.1',
    '1001.02510156.1',
    '1001.02510145.1'
    ]

roomN10 = [
    '102',
    '110',
    '120',
    '124',
    '402',
    '410',
    '420',
    '426',
    '602',
    '610',
    '620',
    '626'

    ]

    # ID11# 阳面
IDS11 = [
    '1001.02520001.1',
    '1001.02520011.1',
    '1001.02520020.1',
    '1001.02520024.1',
    '1001.02520081.1',
    '1001.02520093.1',
    '1001.02520099.1',
    '1001.02520105.1',
    '1001.02520135.1',
    '1001.02520147.1',
    '1001.02520153.1',
    '1001.02520159.1'
    ]

roomS11 = [
    '101',
    '111',
    '119',
    '123',
    '401',
    '411',
    '417',
    '423',
    '601',
    '611',
    '617',
    '623'

    ]
    # 阴面
IDN11 = [
    '1001.02520002.1',
    '1001.02520010.1',
    '1001.02520016.1',
    '1001.02520023.1',
    '1001.02520082.1',
    '1001.02520092.1',
    '1001.02520098.1',
    '1001.02520106.1',
    '1001.02520136.1',
    '1001.02520146.1',
    '1001.02520152.1',
    '1001.02520160.1'
    ]

roomN11 = [
    '102',
    '110',
    '116',
    '122',
    '402',
    '410',
    '416',
    '424',
    '602',
    '610',
    '616',
    '624'
    ]

    # ID12# 阳面
IDS12 = ['1001.02530001.1',
    '1001.02530005.1',
    '1001.02530011.1',
    '1001.02530019.1',
    '1001.02530021.1',
    '1001.02530027.1',
    '1001.02530088.1',
    '1001.02530094.1',
    '1001.02530100.1',
    '1001.02530109.1',
    '1001.02530111.1',
    '1001.02530117.1',
    '1001.02530150.1',
    '1001.02530154.1',
    '1001.02530160.1',
    '1001.02530168.1',
    '1001.02530170.1',
    '1001.02530176.1'

    ]

roomS12 = ['101',
    '105',
    '111',
    '117',
    '119',
    '125',
    '401',
    '405',
    '411',
    '417',
    '419',
    '425',
    '601',
    '605',
    '611',
    '617',
    '619',
    '625',

    ]

    # 阴面
IDN12 = ['1001.02530002.1',
    '1001.02530006.1',
    '1001.02530012.1',
    '1001.02530020.1',
    '1001.02530022.1',
    '1001.02530026.1',
    '1001.02530091.1',
    '1001.02530094.1',
    '1001.02530099.1',
    '1001.02530110.1',
    '1001.02530112.1',
    '1001.02530118.1',
    '1001.02530151.1',
    '1001.02530155.1',
    '1001.02530159.1',
    '1001.02530169.1',
    '1001.02530171.1',
    '1001.02530177.1'

    ]

roomN12 = ['102',
    '106',
    '112',
    '118',
    '120',
    '124',
    '402',
    '406',
    '410',
    '418',
    '420',
    '426',
    '602',
    '606',
    '610',
    '618',
    '620',
    '626'
    ]

    # ID13# 阳面
IDS13 = [
    '1001.02540001.1',
    '1001.02540003.1',
    '1001.02540007.1',
    '1001.02540148.1',
    '1001.02540152.1',
    '1001.02540166.1',
    '1001.02540073.1',
    '1001.02540075.1',
    '1001.02540079.1',
    '1001.02540218.1',
    '1001.02540222.1',
    '1001.02540238.1',
    '1001.02540119.1',
    '1001.02540121.1',
    '1001.02540127.1',
    '1001.02540264.1',
    '1001.02540268.1',
    '1001.02540284.1'

    ]

roomS13 = [
    '101',
    '103',
    '107',
    '123',
    '127',
    '139',
    '401',
    '403',
    '407',
    '423',
    '427',
    '441',
    '机房601',
    '603',
    '607',
    '623',
    '627',
    '641'

    ]

    # 阴面
IDN13 = [
    '1001.02540002.1',
    '1001.02540004.1',
    '1001.02540008.1',
    '1001.02540147.1',
    '1001.02540151.1',
    '1001.02540165.1',
    '1001.02540074.1',
    '1001.02540076.1',
    '1001.02540080.1',
    '1001.02540217.1',
    '1001.02540221.1',
    '1001.02540239.1',
    '1001.02540120.1',
    '1001.02540122.1',
    '1001.02540128.1',
    '1001.02540263.1',
    '1001.02540267.1',
    '1001.02540285.1'
    ]

roomN13 = [

    '102',
    '104',
    '108',
    '122',
    '126',
    '138',
    '402',
    '404',
    '408',
    '422',
    '426',
    '442',
    '602',
    '604',
    '608',
    '622',
    '626',
    '642'

    ]

    # ID14# 阳面
IDS14 = [
    '1001.02550001.1',
    '1001.02550007.1',
    '1001.02550009.1',
    '1001.02550174.1',
    '1001.02550176.1',
    '1001.02550183.1',
    '1001.02550080.1',
    '1001.02550086.1',
    '1001.02550092.1',
    '1001.02550245.1',
    '1001.02550247.1',
    '1001.02550256.1',
    '1001.02550135.1',
    '1001.02550141.1',
    '1001.02550145.1',
    '1001.02550294.1',
    '1001.02550296.1',
    '1001.02550303.1'
    ]

roomS14 = [

    '101',
    '107',
    '109',
    '134',
    '136',
    '141',
    '401',
    '407',
    '409',
    '434',
    '436',
    '443',
    '601',
    '607',
    '609',
    '634',
    '636',
    '643'

    ]

    # 阴面
IDN14 = [
    '1001.02550002.1',
    '1001.02550008.1',
    '1001.02550014.1',
    '1001.02550175.1',
    '1001.02550177.1',
    '1001.02550182.1',
    '1001.02550081.1',
    '1001.02550091.1',
    '1001.02550097.1',
    '1001.02550246.1',
    '1001.02550248.1',
    '1001.02550257.1',
    '1001.02550136.1',
    '1001.02550142.1',
    '1001.02550145.1',
    '1001.02550295.1',
    '1001.02550297.1',
    '1001.02550304.1'
    ]

roomN14 = [

    '102',
    '108',
    '114',
    '135',
    '137',
    '140',
    '402',
    '408',
    '414',
    '435',
    '437',
    '444',
    '602',
    '608',
    '614',
    '635',
    '637',
    '644'

    ]

    # ID15# 阳面
IDS15 = [
    '1001.02560001.1',
    '1001.02560007.1',
    '1001.02560009.1',
    '1001.02560174.1',
    '1001.02560176.1',
    '1001.02560183.1',
    '1001.02560081.1',
    '1001.02560087.1',
    '1001.02560092.1',
    '1001.02560245.1',
    '1001.02560247.1',
    '1001.02560256.1',
    '1001.02560135.1',
    '1001.02560141.1',
    '1001.02560145.1',
    '1001.02560293.1',
    '1001.02560295.1',
    '1001.02560302.1'

    ]

roomS15 = [

    '101',
    '107',
    '109',
    '134',
    '136',
    '141',
    '401',
    '407',
    '409',
    '434',
    '436',
    '443',
    '601',
    '607',
    '609',
    '634',
    '636',
    '643'


    ]

    # 阴面
IDN15 = [
    '1001.02560002.1',
    '1001.02560008.1',
    '1001.02560014.1',
    '1001.02560175.1',
    '1001.02560177.1',
    '1001.02560182.1',
    '1001.02560082.1',
    '1001.02550091.1',
    '1001.02560097.1',
    '1001.02560246.1',
    '1001.02550248.1',
    '1001.02560257.1',
    '1001.02560136.1',
    '1001.02560142.1',
    '1001.02560150.1',
    '1001.02560294.1',
    '1001.02560296.1',
    '1001.02560303.1'
    ]

roomN15 = [

    '102',
    '108',
    '114',
    '135',
    '137',
    '140',
    '402',
    '408',
    '414',
    '435',
    '437',
    '444',
    '602',
    '608',
    '614',
    '635',
    '637',
    '644'

    ]


    # ID16# 阳面
IDS16 = [
    '1001.02570001.1',
    '1001.02570007.1',
    '1001.02570015.1',
    '1001.02570148.1',
    '1001.02570156.1',
    '1001.02570166.1',
    '1001.02570061.1',
    '1001.02570069.1',
    '1001.02570079.1',
    '1001.02570222.1',
    '1001.02570228.1',
    '1001.02570237.1',
    '1001.02570109.1',
    '1001.02570117.1',
    '1001.02570127.1',
    '1001.02570271.1',
    '1001.02570277.1',
    '1001.02570283.1'
    ]

roomS16 = [
    '101',
    '107',
    '115',
    '121',
    '129',
    '135',
    '401',
    '409',
    '417',
    '425',
    '431',
    '437',
    '601',
    '609',
    '617',
    '625',
    '631',
    '637'

    ]

    # 阴面
IDN16 = [
    '1001.02570002.1',
    '1001.02570008.1',
    '1001.02570014.1',
    '1001.02570149.1',
    '1001.02570155.1',
    '1001.02570165.1',
    '1001.02570062.1',
    '1001.02570070.1',
    '1001.02570078.1',
    '1001.02570221.1',
    '1001.02570229.1',
    '1001.02570238.1',
    '1001.02570110.1',
    '1001.02570118.1',
    '1001.02570124.1',
    '1001.02570266.1',
    '1001.02570278.1',
    '1001.02570284.1'
    ]

roomN16 = [
    '102',
    '108',
    '114',
    '122',
    '128',
    '134',
    '402',
    '410',
    '416',
    '424',
    '432',
    '438',
    '602',
    '610',
    '616',
    '624',
    '632',
    '638'
    ]

    # 17# 阳面
IDS17 = [
    '1001.02580145.1',
    '1001.02580163.1',
    '1001.02580011.1',
    '1001.02580021.1',
    '1001.02580221.1',
    '1001.02580239.1',
    '1001.02580080.1',
    '1001.02580092.1',
    '1001.02580273.1',
    '1001.02580291.1',
    '1001.02580128.1',
    '1001.02580138.1'
    ]

roomS17 = [
    '101',
    '117',
    '133',
    '141',
    '401',
    '417',
    '433',
    '443',
    '601',
    '617',
    '633',
    '643'
    ]

    # 阴面
IDN17 = [
    '1001.02580145.1',
    '1001.02580163.1',
    '1001.02580011.1',
    '1001.02580021.1',
    '1001.02580221.1',
    '1001.02580239.1',
    '1001.02580080.1',
    '1001.02580092.1',
    '1001.02580273.1',
    '1001.02580291.1',
    '1001.02580128.1',
    '1001.02580138.1'
    ]

roomN17 = [
    '101',
    '117',
    '133',
    '141',
    '401',
    '417',
    '433',
    '443',
    '601',
    '617',
    '633',
    '643'
    ]

    # ID18# 阳面
IDS18 = [
    '1001.02590001.1',
    '1001.02590007.1',
    '1001.02590011.1',
    '1001.02590015.1',
    '1001.02590020.1',
    '1001.02590026.1',
    '1001.02590092.1',
    '1001.02590098.1',
    '1001.02590102.1',
    '1001.02590110.1',
    '1001.02590114.1',
    '1001.02590120.1',
    '1001.02590154.1',
    '1001.02590160.1',
    '1001.02590166.1',
    '1001.02590172.1',
    '1001.02590176.1',
    '1001.02590184.1'
    ]

roomS18 = [
    '101',
    '107',
    '111',
    '115',
    '119',
    '西南角无人',
    '401',
    '407',
    '411',
    '417',
    '421',
    '427',
    '601',
    '607',
    '611',
    '617',
    '621',
    '627'
    ]

    # 阴面
IDN18 = [
    '1001.02590002.1',
    '1001.02590006.1',
    '1001.02590010.1',
    '1001.02590014.1',
    '1001.02590021.1',
    '1001.02590027.1',
    '1001.02590093.1',
    '1001.02590097.1',
    '1001.02590101.1',
    '1001.02590109.1',
    '1001.02590115.1',
    '1001.02590121.1',
    '1001.02590155.1',
    '1001.02590159.1',
    '1001.02590165.1',
    '1001.02590171.1',
    '1001.02590177.1',
    '1001.02590185.1'

    ]

roomN18 = [
    '102',
    '106',
    '110',
    '114',
    '120',
    '西北角无人',
    '402',
    '406',
    '410',
    '416',
    '422',
    '428',
    '602',
    '606',
    '610',
    '616',
    '622',
    '628'
    ]

    # 19# 阳面
IDS19 = [
    '1001.02600001.1',
    '1001.02600007.1',
    '1001.02600013.1',
    '1001.02600019.1',
    '1001.02600023.1',
    '1001.02600027.1',
    '1001.02600093.1',
    '1001.02600099.1',
    '1001.02600103.1',
    '1001.02600113.1',
    '1001.02600117.1',
    '1001.02600121.1',
    '1001.02600155.1',
    '1001.02600163.1',
    '1001.02600169.1',
    '1001.02600175.1',
    '1001.02600181.1',
    '1001.02600185.1'

    ]

roomS19 = [

    '101',
    '107',
    '113',
    '117',
    '121',
    '125',
    '401',
    '407',
    '411',
    '419',
    '423',
    '427',
    '601',
    '607',
    '613',
    '619',
    '623',
    '627'

    ]

    # 阴面
IDN19 = [
    '1001.02600002.1',
    '1001.02600006.1',
    '1001.02600012.1',
    '1001.02600014.1',
    '1001.02600024.1',
    '1001.02600028.1',
    '1001.02600094.1',
    '1001.02600098.1',
    '1001.02600102.1',
    '1001.02600114.1',
    '1001.02600118.1',
    '1001.02600122.1',
    '1001.02600156.1',
    '1001.02600160.1',
    '1001.02600168.1',
    '1001.02600176.1',
    '1001.02600182.1',
    '1001.02600186.1'

    ]

roomN19 = [

    '102',
    '106',
    '112',
    '114',
    '122',
    '126',
    '402',
    '406',
    '410',
    '420',
    '424',
    '428',
    '602',
    '606',
    '612',
    '620',
    '624',
    '628'
    ]

    # 21# 阳面
IDS21 = [
    '1001.02620001.1',
    '1001.02620007.1',
    '1001.02620015.1',
    '1001.02620020.1',
    '1001.02620022.1',
    '1001.02620026.1',
    '1001.02620073.1',
    '1001.02620077.1',
    '1001.02620081.1',
    '1001.02620089.1',
    '1001.02620130.1',
    '1001.02620134.1',
    '1001.02620148.1',
    '1001.02620152.1',
    '1001.02620156.1',
    '1001.02620166.1',
    '1001.02620172.1',
    '1001.02620176.1'

    ]

roomS21 = [

    '101无人',
    '107',
    '115',
    '119',
    '121',
    '125',
    '401',
    '405',
    '409',
    '417',
    '423',
    '427',
    '601',
    '605',
    '609',
    '617',
    '623',
    '627'

    ]
    # 阴面
IDN21 = [
    '1001.02620002.1',
    '1001.02620008.1',
    '1001.02620014.1',
    '1001.02620019.1',
    '1001.02620021.1',
    '1001.02620027.1',
    '1001.02620074.1',
    '1001.02620078.1',
    '1001.02620082.1',
    '1001.02620088.1',
    '1001.02620129.1',
    '1001.02620135.1',
    '1001.02620149.1',
    '1001.02620153.1',
    '1001.02620157.1',
    '1001.02620165.1',
    '1001.02620171.1',
    '1001.02620177.1'

    ]

roomN21 = [

    '102',
    '108',
    '114',
    '118',
    '120',
    '126',
    '402',
    '406',
    '410',
    '416',
    '422',
    '428',
    '602',
    '606',
    '610',
    '616',
    '622',
    '628'
    ]

    # 22# 阳面
IDS22 = [
    '1001.02630001.1',
    '1001.02630007.1',
    '1001.02630009.1',
    '1001.02630015.1',
    '1001.02630017.1',
    '1001.02630023.1',
    '1001.02630076.1',
    '1001.02630082.1',
    '1001.02630084.1',
    '1001.02630091.1',
    '1001.02630093.1',
    '1001.02630099.1',
    '1001.02630134.1',
    '1001.02630140.1',
    '1001.02630142.1',
    '1001.02630149.1',
    '1001.02630151.1',
    '1001.02630157.1'

    ]

roomS22 = [

    '101',
    '107',
    '109',
    '115',
    '117',
    '121',
    '401',
    '407',
    '409',
    '415',
    '417',
    '423',
    '601',
    '607',
    '609',
    '615',
    '617',
    '623'

    ]
    # 阴面
IDN22 = [
    '1001.02630002.1',
    '1001.02630008.1',
    '1001.02630010.1',
    '1001.02630016.1',
    '1001.02630020.1',
    '1001.02630024.1',
    '1001.02630077.1',
    '1001.02630083.1',
    '1001.02630085.1',
    '1001.02630092.1',
    '1001.02630094.1',
    '1001.02630100.1',
    '1001.02630135.1',
    '1001.02630141.1',
    '1001.02630143.1',
    '1001.02630150.1',
    '1001.02630152.1',
    '1001.02630158.1'

    ]

roomN22 = [

    '102',
    '108',
    '110',
    '116',
    '118',
    '122',
    '402',
    '408',
    '410',
    '416',
    '418',
    '424',
    '602',
    '608',
    '610',
    '616',
    '618',
    '624'

    ]

    # 23# 阳面
IDS23 = [
    '1001.02640207.1',
    '1001.02640219.1',
    '1001.02640221.1',
    '1001.02640241.1',
    '1001.02640243.1',
    '1001.02640251.1',
    '1001.02640042.1',
    '1001.02640052.1',
    '1001.02640055.1',
    '1001.02640116.1',
    '1001.02640118.1',
    '1001.02640149.1',
    '1001.02640086.1',
    '1001.02640097.1',
    '1001.02640099.1',
    '1001.02640156.1',
    '1001.02640158.1',
    '1001.02640196.1'

    ]

roomS23 = [
    '101',
    '112',
    '114',
    '126',
    '128',
    '136',
    '401',
    '411',
    '413',
    '425',
    '427',
    '439',
    '601',
    '611',
    '613',
    '625',
    '627',
    '639'

    ]
    # 阴面
IDN23 = [
    '1001.02640208.1',
    '1001.02640218.1',
    '1001.02640220.1',
    '1001.02640235.1',
    '1001.02640236.1',
    '1001.02640248.1',
    '1001.02640043.1',
    '1001.02640053.1',
    '1001.02640056.1',
    '1001.02640117.1',
    '1001.02640119.1',
    '1001.02640150.1',
    '1001.02640087.1',
    '1001.02640098.1',
    '1001.02640100.1',
    '1001.02640157.1',
    '1001.02640159.1',
    '1001.02640197.1'

    ]

roomN23 = [
    '102',
    '111',
    '113',
    '125',
    '127',
    '137',
    '402',
    '412',
    '414',
    '426',
    '428',
    '440',
    '602',
    '612',
    '614',
    '626',
    '628',
    '640'
    ]

    # 24# 阳面
IDS24 = [
    '1001.02650001.1',
    '1001.02650013.1',
    '1001.02650022.1',
    '1001.02650034.1',
    '1001.02650109.1',
    '1001.02650121.1',
    '1001.02650128.1',
    '1001.02650142.1',
    '1001.02650181.1',
    '1001.02650193.1',
    '1001.02650200.1',
    '1001.02650214.1'

    ]

roomS24 = [
    '101',
    '113',
    '119',
    '131',
    '401',
    '413',
    '419',
    '433',
    '601',
    '613',
    '619',
    '633'

    ]
    # 阴面
IDN24 = [
    '1001.02650004.1',
    '1001.02650012.1',
    '1001.02650021.1',
    '1001.02650033.1',
    '1001.02650110.1',
    '1001.02650120.1',
    '1001.02650127.1',
    '1001.02650143.1',
    '1001.02650182.1',
    '1001.02650192.1',
    '1001.02650199.1',
    '1001.02650215.1'

    ]

roomN24 = [
    '104',
    '112',
    '118',
    '130',
    '402',
    '412',
    '418',
    '434',
    '602',
    '612',
    '618',
    '634'
    ]

    # 25# 阳面
IDS25 = [
    '1001.02660001.1',
    '1001.02660009.1',
    '1001.02660011.1',
    '1001.02660017.1',
    '1001.02660020.1',
    '1001.02660028.1',
    '1001.02660109.1',
    '1001.02660117.1',
    '1001.02660119.1',
    '1001.02660125.1',
    '1001.02660128.1',
    '1001.02660138.1',
    '1001.02660181.1',
    '1001.02660189.1',
    '1001.02660191.1',
    '1001.02660197.1',
    '1001.02660200.1',
    '1001.02660210.1'

    ]

roomS25 = [
    '101',
    '109',
    '111',
    '117',
    '119',
    '127',
    '401',
    '409',
    '411',
    '417',
    '419',
    '429',
    '601',
    '609',
    '611',
    '617',
    '619',
    '629'

    ]
    # 阴面
IDN25 = [
    '1001.02660002.1',
    '1001.02660010.1',
    '1001.02660012.1',
    '1001.02660019.1',
    '1001.02660021.1',
    '1001.02660029.1',
    '1001.02660110.1',
    '1001.02660118.1',
    '1001.02660120.1',
    '1001.02660127.1',
    '1001.02660129.1',
    '1001.02660139.1',
    '1001.02660182.1',
    '1001.02660190.1',
    '1001.02660192.1',
    '1001.02660199.1',
    '1001.02660201.1',
    '1001.02660211.1'

    ]

roomN25 = [
    '102',
    '110',
    '112',
    '118',
    '120',
    '128',
    '402',
    '410',
    '412',
    '418',
    '420',
    '430',
    '602',
    '610',
    '612',
    '618',
    '620',
    '630'

    ]

    # 26# 阳面
IDS26 = [
    '1001.02670001.1',
    '1001.02670011.1',
    '1001.02670013.1',
    '1001.02670022.1',
    '1001.02670024.1',
    '1001.02670034.1',
    '1001.02670113.1',
    '1001.02670123.1',
    '1001.02670125.1',
    '1001.02670128.1',
    '1001.02670130.1',
    '1001.02670142.1',
    '1001.02670199.1',
    '1001.02670209.1',
    '1001.02670211.1',
    '1001.02670236.1',
    '1001.02670238.1',
    '1001.02670250.1'

    ]
roomS26 = [
    '101',
    '111',
    '113',
    '121',
    '123',
    '133',
    '401',
    '411',
    '413',
    '421',
    '423',
    '435',
    '601',
    '611',
    '613',
    '621',
    '623',
    '635'

    ]
    # 阴面
IDN26 = [
    '1001.02670002.1',
    '1001.02670012.1',
    '1001.02660014.1',
    '1001.02670023.1',
    '1001.02670025.1',
    '1001.02670033.1',
    '1001.02670114.1',
    '1001.02670124.1',
    '1001.02670145.1',
    '1001.02670129.1',
    '1001.02670131.1',
    '1001.02670143.1',
    '1001.02670200.1',
    '1001.02670210.1',
    '1001.02670212.1',
    '1001.02670237.1',
    '1001.02670239.1',
    '1001.02670251.1'

    ]
roomN26 = [
    '102',
    '112',
    '114',
    '122',
    '124',
    '132',
    '402',
    '412',
    '414',
    '422',
    '424',
    '436',
    '602',
    '612',
    '614',
    '622',
    '624',
    '636'

    ]

    # 27# 阳面
IDS27 = [
    '1001.02680001.1',
    '1001.02680005.1',
    '1001.02680007.1',
    '1001.02680011.1',
    '1001.02680013.1',
    '1001.02680019.1',
    '1001.02680062.1',
    '1001.02680066.1',
    '1001.02680068.1',
    '1001.02680073.1',
    '1001.02680075.1',
    '1001.02680081.1',
    '1001.02680109.1',
    '1001.02680113.1',
    '1001.02680115.1',
    '1001.02680119.1',
    '1001.02680121.1',
    '1001.02680128.1'

    ]

roomS27 = [
    '101',
    '105',
    '107',
    '111',
    '113',
    '118',
    '401',
    '405',
    '407',
    '411',
    '413',
    '419',
    '601',
    '605',
    '607',
    '611',
    '613',
    '619'

    ]
    # 阴面
IDN27 = [
    '1001.02680002.1',
    '1001.02680006.1',
    '1001.02680010.1',
    '1001.02680012.1',
    '1001.02680014.1',
    '1001.02680017.1',
    '1001.02680063.1',
    '1001.02680067.1',
    '1001.02680069.1',
    '1001.02680074.1',
    '1001.02680076.1',
    '1001.02680082.1',
    '1001.02680110.1',
    '1001.02680114.1',
    '1001.02680116.1',
    '1001.02680120.1',
    '1001.02680122.1',
    '1001.02680129.1'

    ]

roomN27 = [
    '102',
    '106',
    '110',
    '112',
    '114',
    '117',
    '402',
    '406',
    '408',
    '412',
    '414',
    '420',
    '602',
    '606',
    '608',
    '612',
    '614',
    '620'

    ]

n  = -1
rang = ['6','7','8','9','10','11','12','13','14','15','16','17','18','19','21','22','23','24','25','26','27']

if __name__ == '__main__':
    wb = workbook.Workbook()
    ws = wb.active

#  For sunny side
    for i in rang:
        ID1 = locals()['IDS' + i ]   # A way to get variable name regularly
        room = locals()['roomS' + i]
        FileName = i + "#阳面"
        print(i)
        for id in ID1:
            n += 1
            ws.append(["房间号"])
            ws.append([room[n]])
            ws.append(["日期","用量"])
            getHtml(id,SD,END)
        n = -1
        wb.save(FileName + ".xlsx")
# For night side
    for i in rang:
        ID2 = locals()['IDN' + i ]   # A way to get variable name regularly
        room = locals()['roomN' + i]
        FileName = i + "#阴面"
        print(i)
        for id in ID2:
            n += 1
            ws.append(["房间号"])
            ws.append([room[n]])
            ws.append(["日期","用量"])
            getHtml(id,SD,END)
        n = -1
        wb.save(FileName + ".xlsx")
