import IEC
import os


zipcodes = {'00841': 0, '00840': 0, '00843': 0, '00987': 1,
            '00846': 1, '00819': 1, '00707': 1,
            '00720': 1, '00721': 1, '00723': 1, '00725': 1,
            '00928': 1, '00727': 1, '00683': 1, '00729': 1,
            '00924': 1, '00680': 1, '00687': 1, '00923': 1, '00685': 1,
            '00921': 1, '00603': 1, '00601': 1, '00607': 1, '00606': 1,
            '00604': 1, '00609': 1, '00850': 1, '00851': 1, '00692': 1,
            '00693': 1, '00698': 1, '00918': 1, '00913': 1, '00751': 1,
            '00911': 1, '00917': 1, '00915': 1, '00914': 1, '00754': 1,
            '00757': 1, '00618': 1, '00619': 1, '00753': 1, '00615': 1,
            '00617': 1, '00610': 1, '00611': 1, '00612': 1, '00613': 1,
            '00907': 1, '00759': 1, '00864': 1, '00909': 1, '00861': 1,
            '00758': 1, '00728': 1, '00465': 1, '00927': 1, '00744': 1,
            '00745': 1, '00740': 1, '00741': 1, '00661': 1, '00660': 1,
            '00986': 1, '00662': 1, '00665': 1, '00664': 1, '00667': 1,
            '00666': 1, '00920': 1, '00802': 1, '00970': 1, '00977': 1,
            '00976': 1, '00979': 1, '00748': 1, '00773': 1, '00772': 1,
            '00678': 1, '00777': 1, '00775': 1, '00672': 1, '00673': 1,
            '00670': 1, '00671': 1, '00674': 1, '00982': 1, '00962': 1,
            '00963': 1, '00961': 1, '00949': 1, '00968': 1, '00969': 1,
            '00888': 1, '00646': 1, '00645': 1, '00644': 1, '00643': 1,
            '00642': 1, '00641': 1, '00640': 1, '00805': 1, '00804': 1,
            '00801': 1, '00803': 1, '00648': 1, '00768': 1, '00764': 1,
            '00765': 1, '00767': 1, '00760': 1, '00761': 1, '00763': 1,
            '00117': 1, '00791': 1, '00792': 1, '00795': 1, '00794': 1,
            '00926': 1, '00957': 1, '00956': 1, '00954': 1, '00953': 1,
            '00771': 1, '00951': 1, '00950': 1, '00650': 1, '00340': 1,
            '00654': 1, '00655': 1, '00656': 1, '00657': 1, '00658': 1,
            '00659': 1, '00812': 1, '00813': 1, '00719': 1, '00718': 1,
            '00890': 1, '00711': 1, '00717': 1, '00929': 1, '00782': 1, '00783': 1,
            '00780': 1, '00778': 1, '00784': 1, '00785': 1, '00983': 1, '00823': 1,
            '00822': 1, '00821': 1, '00820': 1, '00824': 1, '00625': 1, '00708': 1,
            '00626': 1, '00959': 1, '00701': 1, '00629': 1, '00628': 1, '00704': 1,
            '00705': 1, '00834': 1, '00830': 1, '00831': 1, '00832': 1, '00627': 1,
            '00737': 1, '00736': 1, '00735': 1, '00709': 1, '00925': 1, '00732': 1,
            '00731': 1, '00936': 1, '00739': 1, '00738': 1, '00637': 1, '00635': 1,
            '00632': 1, '00633': 1, '00630': 1, '00638': 1, '00639': 1, '00984': 1,
            '00985': 1}
##
##proxy_info = {'user':'mshomphe',
##              'password':'SrV@16890',
##              'host':'SIMPROXY',
##              'port':'80'}
##
###os.environ['HTTP_PROXY'] = 'http://%(user)s:%(password)s@%(host)s:%(port)s' % proxy_info
##os.environ['HTTP_PROXY'] = 'http://%(host)s' % proxy_info
###print 'Proxy is: %s' % os.environ['HTTP_PROXY']
error_str = 'The ZIP Code you entered could not be found in our database.'.lower()
##data = {'zipcode':''}
####
####proxy = urllib2.ProxyHandler({'http':'http://SIMPROXY'})
####proxy_auth_handler = urllib2.ProxyBasicAuthHandler()
####proxy_auth_handler.add_password('', '', 'mshomphe  ', 'bibble')
####my_request = urllib2.Request('http://zip4.usps.com/zip4/zip_responseA.jsp')
####opener = urllib2.build_opener(proxy, proxy_auth_handler)
####urllib2.install_opener(opener)
##url = 'http://zip4.usps.com/zip4/zip_responseA.jsp?%s'
##test_url = "http://www.python.org/index.html"
##for z in zipcodes.keys():
##    data['zipcode'] = z
##    my_request = url % urllib.urlencode(data)
##    #my_request.add_data(urllib.urlencode(data))
##    print 'Opening URL: %s' %my_request
####    print 'Opening URL: %s' % my_request.get_full_url()
####    print 'Data are: %s' % my_request.get_data()
####    print 'method is: %s' % my_request.get_method()
##    handle = urllib2.urlopen(test_url)
##    txt = handle.read().lower()
##    handle.close()
##    print "Text: ", txt, '-'*80
##    if txt.find(error_str) >= 0:
##        print 'BAD ZIPCODE'
##	zipcodes[z] = 0
##    else:
##        zipcodes[z] = 1
##
##for invalid_zip in zipcodes.keys():
##    if zipcodes[invalid_zip] == 0:
##        print invalid_zip

##conn = httplib.HTTPSConnection("%(user)s:%(password)s@%(host)s"%proxy_info,80)
##conn.request("GET", "www.cnn.com", "/")
##conn.endheaders()
##r = conn.getresponse() 
##print r.status, r.reason 
##print r.msg 
##while 1: 
##     data = r.read(1024) 
##     if len(data) < 1024: break 
##     print data

##proxy_info = {'user':'us3r',
##              'password':'p@ssword',
##              'host':'MY_PROXY',
##              'port':'80'}
##os.environ['HTTP_PROXY'] = 'http://%(user)s:%(password)s@%(host)s:%(port)s' % proxy_info
##test_url = "http://www.python.org/index.html"
###handle = urllib2.urlopen(test_url)
##handle = urllib.urlopen(test_url)
##txt = handle.read().lower()
##handle.close()
##print "Text: " 
##print txt
url = 'http://zip4.usps.com/zip4/zip_responseA.jsp?zipcode=%s'
explorer = IEC.IEController() 

for z in zipcodes.keys():
    explorer.Navigate(url%z)
    if explorer.GetDocumentText().lower().find(error_str) > -1:
        zipcodes[z] = 0
    else:
        zipcodes[z] = 1
explorer.CloseWindow()
zips = zipcodes.keys()
zips.sort()
for invalid_zip in zips:
    if zipcodes[invalid_zip] == 0:
        print "Bad:",invalid_zip
    else:
        print "Good:",invalid_zip
