import xlrd
L_F=[15, 30,45,60,75,90,105,120,135,150,165,195,225,255,285]
F_T=[14,27,31,34,35,40,42,48,50,53,56,66,85,121,123,139,146,151,156,168,184,197,200,202,
205,206,209,210,211,215,218,227,245,246,247,252,256,269,275,286,288,291,293]
F_F=[17,20,54,65,75,83,112,113,115,164,169,177,185,196,199,220,257,258,272,276]
K_T=[96]
K_F=[30,39,71,89,124,129,134,138,142,148,160,170,171,180,183,217,234,267,272,296,316,322,368,370,372,373,375,386,394]
Hs_T=[23,29,43,62,72,108,114,125,161,189,273]
Hs_F=[2,3,7,9,18,51,55,63,68,103,130,155,163,175,188,190,192,230,243,274]
D_T=[5,32,41,43,52,67,68,104,130,138,142,158,159,182,189,193,236,259,288,290]
D_F=[2,8,9,18,30,36,39,46,51,57,58,64,80,88,89,95,98,107,122,131,145,152,153,154,155,160,178,191,207,208,233,241,242,248,263,270,271,272,285,296]
Hy_T=[10,23,32,43,44,47,76,114,179,186,189,238,253]
Hy_F=[2,3,6,7,8,9,12,26,30,51,55,71,89,93,103,107,109,124,128,129,136,137,141,147,153,160,162,163,170,172,174,175,180,188,190,192,201,213,230,234,243,265,267,274,279,289,292]
Pd_T=[16,21,24,32,33,35,38,42,61,67,84,94,102,106,110,118,127,215,216,224,239,244,245,284]
Pd_F=[8,20,37,82,91,96,107,134,137,141,155,170,171,173,180,183,201,231,235,237,248,267,287,289,294,296]
Mfm_T=[4,25,69,70,74,77,78,87,92,126,132,134,140,149,179,187,203,204,217,226,231,239,261,278,282,295,297,299]
Mfm_F=[1,19,26,28,79,80,81,89,99,112,115,116,117,120,133,144,176,198,213,214,219,221,223,229,231,249,254,260,262,264,280,283,297,300]
Mff_T=[4,25,70,74,77,78,87,92,126,132,133,134,140,149,187,203,204,217,226,239,261,278,282,235,299]
Mff_F=[1,19,26,28,69,79,80,81,99,112,115,116,117,120,144,176,179,198,213,214,219,221,223,229,231,249,254,260,262,264,280,283,297,300]
Pa_T=[16,24,27,35,110,121,123,127,151,157,158,202,275,284,291,293,299,305,314,317,326,338,341,364,365]
Pa_F=[93,107,109,111,117,124,268,281,294,313,316,319,327,347,348]
Pt_T=[10,15,22,32,41,67,76,86,94,102,106,142,159,182,189,217,238,266,301,304,321,336,337,340,342,343,344,346,349,351,352,356,357,358,359,360,361,362,366]
Pt_F=[3,8,36,122,152,164,178,329,353]
Sc_T=[15,22,40,41,47,52,76,97,104,121,156,157,159,168,179,182,184,202,210,212,238,241,251,259,266,273,282,291,297,301,303,307,308,311,312,
      315,320,323,324,325,328,331,332,333,334,335,339,341,345,349,350,352,354,355,356,360,363,364,366]
Sc_F=[17,65,103,119,177,178,187,192,196,220,276,281,302,306,309,310,318,322,330]
Ma_T=[11,13,21,22,59,64,73,97,100,109,127,134,143,156,157,167,181,194,212,222,226,228,232,233,233,240,250,251,263,266,253,271,277,279,298]
Ma_F=[101,105,111,119,120,148,166,171,180,267,289]
Si_T=[32,67,82,111,117,124,138,147,171,172,180,201,236,267,278,292,304,316,321,332,336,342,357,360,370,373,376,378,379,385,389,393,398,399]
Si_F=[25,33,57,91,99,110,126,143,193,208,229,231,254,262,281,296,309,353,359,367,371,374,377,380,381,382,383,384,387,388,390,391,392,395,396,397]
L=0
F=0
K=0
Hs=0
D=0
Hy=0
Pd=0
Mfm=0
Mff=0
Pa=0
Pt=0
Sc=0
Ma=0
Si=0
dic={}
new_wb = xlrd.open_workbook("?????????.xls", formatting_info=True)
new_sh = new_wb.sheet_by_index(0)
for j in range(3,70):
    for i in range(0,11,2):
        key=new_sh.cell_value(j,i)
        value=new_sh.cell_value(j,i+1)
        dic[key]=value
# print(dic.items())
for i in dic:
    if dic[i]=='1' or dic[i]=='1.0':
        i=int(i)
        if i in K_T:
            K+=1
        if i in F_T:
            F+=1
        if i in Hs_T:
            Hs+=1
        if i in D_T:
            D+=1
        if i in Hy_T:
            Hy+=1
        if i in Pd_T:
            Pd+=1
        if i in Mfm_T:
            Mfm+=1
        if i in Mff_T:
            Mff+=1
        if i in Pa_T:
            Pa+=1
        if i in Pt_T:
            Pt+=1
        if i in Sc_T:
            Sc+=1
        if i in Ma_T:
            Ma+=1
        if i in Si_T:
            Si+=1
    elif dic[i]=='0':
        i=int(i)
        if i in L_F:
            L+=1
        if i in K_F:
            K+=1
        if i in F_F:
            F+=1
        if i in Hs_F:
            Hs+=1
        if i in D_F:
            D+=1
        if i in Hy_F:
            Hy+=1
        if i in Pd_F:
            Pd+=1
        if i in Mfm_F:
            Mfm+=1
        if i in Mff_F:
            Mff+=1
        if i in Pa_F:
            Pa+=1
        if i in Pt_F:
            Pt+=1
        if i in Sc_F:
            Sc+=1
        if i in Ma_F:
            Ma+=1
        if i in Si_F:
            Si+=1
print("L????????????:",L)
print("F????????????:",F)
print("K????????????:",K)
print("Hs????????????:",Hs)
print("D????????????:",D)
print("Hy????????????:",Hy)
print("Pd????????????:",Pd)
print("Mf-m????????????:",Mfm)
print("Mf-f????????????:",Mfm)
print("Pa????????????:",Pa)
print("Pt????????????:",Pt)
print("Sc????????????:",Sc)
print("Ma????????????:",Ma)
print("Si????????????:",Si)