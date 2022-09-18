import xlwings as xw
ws = xw.Book("C:\\ESS\\input.xlsx").sheets['Sheet1']
v1=int(ws.range("A2").value)
v2=str(ws.range("B2").value)
v3=int(ws.range("C2").value)
import os.path
a=v2
enodeB=v1
save_path="C:\\ESS"
completeName = os.path.join(save_path, f"{a}"+".txt")
f=open(completeName,"w+")
l=["A","B","C","D","E","F"]
i=71
j=0
for cell in l:
    f.write(f"gs+\ncrn GNBCUCPFunction=1,NRCellCU=N07-{a}-1{cell},EUtranCellRelation=L07-{a}-1{cell}\ncellIndividualOffset 0\nessEnabled true\nisHoAllowed true\nisRemoveAllowed false\nneighborCellRef GNBCUCPFunction=1,EUtraNetwork=1,ExternalENodeBFunction=5205-{enodeB},ExternalEUtranCell=5205-{enodeB}-{i}\nuserLabel\nend\ngs-\n")
    i=i+1
    j=j+1
    if j==v3:
        break
f.close()
