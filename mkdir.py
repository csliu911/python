'''
Created on 2019年7月18日

@author: CM3977-18
'''
import os
import shutil

def mkdir(path):
    folder = os.path.exists(path)
    if not folder:
        os.makedirs(path)
        print("--- new folder... ---")
        print("--- OK ---")
    else:
        print("--- There is this folder! ---")
        
'''
@ make F1K directory
'''
msn = ['adc','can','canV2','cortst','dio','eth','fls','flstst','fr','gpt','icu','lin','mcu','port','pwm','ramtst','spi','wdg']
MSN = ['ADC','CAN','CAN','CORTST','DIO','ETH','FLS','FLSTST','FR','GPT','ICU','LIN','MCU','PORT','PWM','RAMTST','SPI','WDG']
for i in range(len(msn)):
    if(msn[i] == 'can' or msn[i] == 'eth' or msn[i] == 'fr'):
        path_dst = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1KM_Ver4.05.00_Ver42.05.00_F1KH_Ver42.05.00_ASILB"
    elif(msn[i] == 'canV2'):
        path_dst = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1K_Ver4.05.00_Ver42.05.00_ASILB"
    else:
        path_dst = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1K_F1KM_Ver4.05.00_Ver42.05.00_F1KH_Ver42.05.00_ASILB"
#    print(msn[i])
#    print(path_dst)
    mkdir(path_dst)

'''
@ copy safety analysis file from 4.04.00 to 4.05.00
'''
for i in range(len(msn)):
#folder and file names are different for can and canv2 and fr module
#    print(msn[i])
#    print(MSN[i])
    if(msn[i] == 'can' or msn[i] == 'eth' or msn[i] == 'fr'):
        srcfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1KM_Ver4.04.00_Ver42.04.00_F1KH_Ver42.04.00_ASILB\\F1KM_Ver4.04.00_Ver42.04.00_F1KH_Ver42.04.00_SafetyAnalysis_"+MSN[i]+".xlsx"
        dstfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1KM_Ver4.05.00_Ver42.05.00_F1KH_Ver42.05.00_ASILB\\F1KM_Ver4.05.00_Ver42.05.00_F1KH_Ver42.05.00_SafetyAnalysis_"+MSN[i]+".xlsx"
    elif(msn[i] == 'canV2'):        
        srcfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1K_Ver4.04.00_Ver42.04.00_ASILB\\F1K_Ver4.04.00_Ver42.04.00_SafetyAnalysis_"+MSN[i]+".xlsx"
        dstfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1K_Ver4.05.00_Ver42.05.00_ASILB\\F1K_Ver4.05.00_Ver42.05.00_SafetyAnalysis_"+MSN[i]+".xlsx"
    else:
        srcfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1K_F1KM_Ver4.04.00_Ver42.04.00_F1KH_Ver42.04.00_ASILB\\F1K_F1KM_Ver4.04.00_Ver42.04.00_F1KH_Ver42.04.00_SafetyAnalysis_"+MSN[i]+".xlsx"
        dstfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1K_F1KM_Ver4.05.00_Ver42.05.00_F1KH_Ver42.05.00_ASILB\\F1K_F1KM_Ver4.05.00_Ver42.05.00_F1KH_Ver42.05.00_SafetyAnalysis_"+MSN[i]+".xlsx"

    src = os.path.exists(srcfile)
    if not src:
        print("src not exist",srcfile)

    dst = os.path.exists(dstfile)
    if not dst:
        shutil.copy(srcfile, dstfile)
        print("dstfile copied from srcfile")
    else:
        print("dst already exist",dstfile)
'''
@ copy DFA CC files from 4.04.00 to 4.05.00
'''
for i in range(len(msn)):
#    print(msn[i])
#    print(MSN[i])
#folder and file names are different for can and canv2 and fr module
    if(msn[i] == 'can' or msn[i] == 'eth' or msn[i] == 'fr'):
        srcfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1KM_Ver4.04.00_Ver42.04.00_F1KH_Ver42.04.00_ASILB\\F1KM_Ver4.04.00_Ver42.04.00_F1KH_Ver42.04.00_DFA_CC_"+MSN[i]+".xlsm"
        dstfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1KM_Ver4.05.00_Ver42.05.00_F1KH_Ver42.05.00_ASILB\\F1KM_Ver4.05.00_Ver42.05.00_F1KH_Ver42.05.00_DFA_CC_"+MSN[i]+".xlsm"
    elif(msn[i] == 'canV2'):
        srcfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1K_Ver4.04.00_Ver42.04.00_ASILB\\F1K_Ver4.04.00_Ver42.04.00_DFA_CC_"+MSN[i]+".xlsm"
        dstfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1K_Ver4.05.00_Ver42.05.00_ASILB\\F1K_Ver4.05.00_Ver42.05.00_DFA_CC_"+MSN[i]+".xlsm"
    else:
        srcfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1K_F1KM_Ver4.04.00_Ver42.04.00_F1KH_Ver42.04.00_ASILB\\F1K_F1KM_Ver4.04.00_Ver42.04.00_F1KH_Ver42.04.00_DFA_CC_"+MSN[i]+".xlsm"
        dstfile = "U:\\internal\\X1X\\F1x\\modules\\" + msn[i] + "\\docs\\safety_analysis\\F1K_F1KM_Ver4.05.00_Ver42.05.00_F1KH_Ver42.05.00_ASILB\\F1K_F1KM_Ver4.05.00_Ver42.05.00_F1KH_Ver42.05.00_DFA_CC_"+MSN[i]+".xlsm"

    src = os.path.exists(srcfile)
    if not src:
        print("src not exist",srcfile)

    dst = os.path.exists(dstfile)
    if not dst:
        shutil.copy(srcfile, dstfile)
        print("dstfile copied from srcfile")
    else:
        print("dst already exist",dstfile)

print("---done---")

# end of file   