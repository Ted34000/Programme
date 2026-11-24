import json, random
random.seed(7)
cc, la = "c-courant", "c-livret"
d = {"version":1,"exemple":True,
 "params":{"alerteJours":3,"horizonMois":6,"moisMoyenne":3},
 "comptes":[
  {"id":cc,"nom":"Compte courant","banque":"Ma banque","type":"courant","seuil":200,"utilite":"Reçoit les revenus, paie le loyer, les abonnements et les courses"},
  {"id":la,"nom":"Livret A","banque":"Ma banque","type":"epargne","seuil":None,"utilite":"Épargne de précaution, alimentée chaque début de mois"}],
 "operations":[],"releves":[],"depenses":[],
 "categories":[
  {"id":"k-alim","nom":"Alimentation","inclure":True,"compteId":cc,"fixe":None},
  {"id":"k-carb","nom":"Carburant","inclure":True,"compteId":cc,"fixe":None},
  {"id":"k-lois","nom":"Loisirs et sorties","inclure":True,"compteId":cc,"fixe":None},
  {"id":"k-sante","nom":"Santé","inclure":False,"compteId":cc,"fixe":None}],
 "objectifs":[
  {"id":"o1","type":"epargne","libelle":"Épargne de précaution","compteId":la,"cible":5000,"echeance":"2027-12-31"},
  {"id":"o2","type":"seuil","libelle":"Ne jamais passer sous 200 € sur le compte courant","scope":cc,"minimum":200},
  {"id":"o3","type":"plafond","libelle":"Courses à 330 € par mois maximum","categorie":"Alimentation","plafond":330},
  {"id":"o4","type":"taux","libelle":"Épargner 10 % de mes revenus","taux":10}]}
ops = [
 ("Revenu net (salaire)","credit",1850,"mensuel","2026-01-28",cc,"Revenus"),
 ("Loyer","debit",620,"mensuel","2026-01-05",cc,"Logement"),
 ("Électricité","debit",58,"mensuel","2026-01-10",cc,"Logement"),
 ("Box internet","debit",29.99,"mensuel","2026-01-08",cc,"Abonnements"),
 ("Forfait mobile","debit",15.99,"mensuel","2026-01-12",cc,"Abonnements"),
 ("Assurance auto","debit",42,"mensuel","2026-01-15",cc,"Assurances"),
 ("Abonnement streaming","debit",13.49,"mensuel","2026-01-20",cc,"Abonnements"),
 ("Assurance habitation","debit",148,"annuel","2026-03-02",cc,"Assurances"),
 ("Taxe foncière / habitation (exemple)","debit",210,"annuel","2026-10-15",cc,"Impôts"),
]
for i,(l,s,m,f,deb,c,cat) in enumerate(ops):
  d["operations"].append({"id":f"op{i}","libelle":l,"sens":s,"montant":m,"freq":f,"debut":deb,"fin":"","compteId":c,"compteDest":"","categorie":cat,"actif":True})
d["operations"].append({"id":"opv","libelle":"Virement épargne","sens":"virement","montant":150,"freq":"mensuel","debut":"2026-01-01","fin":"","compteId":cc,"compteDest":la,"categorie":"Épargne","actif":True})
d["releves"] = [
 {"id":"r1","compteId":cc,"date":"2026-08-31","solde":1180},
 {"id":"r2","compteId":la,"date":"2026-08-31","solde":2950},
 {"id":"r3","compteId":cc,"date":"2026-09-30","solde":1236.4},
 {"id":"r4","compteId":la,"date":"2026-09-30","solde":3100},
 {"id":"r5","compteId":cc,"date":"2026-10-31","solde":1092.15},
 {"id":"r6","compteId":la,"date":"2026-10-31","solde":3250},
 {"id":"r7","compteId":cc,"date":"2026-11-23","solde":512.3},
 {"id":"r8","compteId":la,"date":"2026-11-23","solde":3400}]
n=0
for mth,last in [("2026-08",31),("2026-09",30),("2026-10",31),("2026-11",23)]:
  for day in [3,10,17,24,30]:
    if day<=last: n+=1; d["depenses"].append({"id":f"e{n}","date":f"{mth}-{day:02d}","categorie":"Alimentation","montant":round(random.uniform(55,85),2),"compteId":cc,"note":"Courses"})
  for day in [6,21]:
    if day<=last: n+=1; d["depenses"].append({"id":f"e{n}","date":f"{mth}-{day:02d}","categorie":"Carburant","montant":round(random.uniform(52,68),2),"compteId":cc,"note":"Plein"})
  for day in [13]:
    n+=1; d["depenses"].append({"id":f"e{n}","date":f"{mth}-{day:02d}","categorie":"Loisirs et sorties","montant":round(random.uniform(35,90),2),"compteId":cc,"note":"Sortie"})
n+=1; d["depenses"].append({"id":f"e{n}","date":"2026-10-18","categorie":"Santé","montant":25,"compteId":cc,"note":"Pharmacie"})
json.dump(d, open("seed.json","w"), ensure_ascii=False)
print(len(d["depenses"]))
# recalcul des soldes du compte courant pour qu'ils collent aux flux (petits écarts réalistes)
import datetime as dt
def occ(o, a, b):
  d0=dt.date.fromisoformat(o["debut"]); out=[]
  step={"mensuel":1,"annuel":12}[o["freq"]]
  k=0
  while True:
    m=d0.month-1+k*step; y=d0.year+m//12; m=m%12+1
    d=dt.date(y,m,d0.day)
    if d>b: break
    if d>a: out.append(d)
    k+=1
  return out
bal=1180.0; prev=dt.date(2026,8,31); noise=[-12.4, 8.15, -4.6]
for i,(rid,date) in enumerate([("r3","2026-09-30"),("r5","2026-10-31"),("r7","2026-11-23")]):
  dd=dt.date.fromisoformat(date)
  for o in d["operations"]:
    if o["compteId"]!=cc: continue
    n=len(occ(o,prev,dd)); s = o["montant"]*(1 if o["sens"]=="credit" else -1)
    bal+=n*s
  bal-=sum(e["montant"] for e in d["depenses"] if prev<dt.date.fromisoformat(e["date"])<=dd)
  bal=round(bal+noise[i],2); prev=dd
  next(r for r in d["releves"] if r["id"]==rid)["solde"]=bal
print([ (r["date"],r["solde"]) for r in d["releves"] if r["compteId"]==cc])
json.dump(d, open("seed.json","w"), ensure_ascii=False)
