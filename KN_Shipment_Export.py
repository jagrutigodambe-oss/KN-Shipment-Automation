
# KN_Shipment_Export.py
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
import pandas as pd
import os

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

TOKEN_URL="https://portal.api.kuehne-nagel.com/oauth2/token"
SEARCH_URL="https://gateway.api.kuehne-nagel.com/track-trace/shipment/v2/shipments-search"
DETAIL_URL="https://gateway.api.kuehne-nagel.com/track-trace/shipment/v2/shipments/shipment-id%3A{}"

def auth():
    s=requests.Session()
    r=s.post(TOKEN_URL,data={"grant_type":"client_credentials"},auth=(CLIENT_ID,CLIENT_SECRET))
    r.raise_for_status()
    s.headers.update({
        "Authorization":"Bearer "+r.json()["access_token"],
        "Accept":"application/json",
        "Content-Type":"application/json"})
    return s

def get_shipments(s):
    out=[];page=0
    while True:
        r=s.post(SEARCH_URL,json={"pagination":{"page":page,"size":500}})
        r.raise_for_status()
        d=r.json(); out.extend(d.get("shipments",[]))
        print(f"Page {page+1}: {len(d.get('shipments',[]))}")
        if not d["page"]["hasNext"]: break
        page+=1
    return out

def fetch(s, sid):
    r=s.get(DETAIL_URL.format(sid))
    if r.ok: return r.json()
    return None

def build(details):
    containers=[];summary=[];routes=[];milestones=[];parties=[]
    for sh in details:
        sid=sh.get("shipmentId");trk=sh.get("trackingNumber")
        cl=sh.get("freightInfo",{}).get("containerInfo",{}).get("containers",[])
        for c in cl:
            row=pd.json_normalize(c).to_dict("records")[0]
            row["Shipment ID"]=sid; row["Tracking Number"]=trk; containers.append(row)
            if not c.get("isVirtual",False):
                eta=None
                mlist=c.get("routing",{}).get("milestoneInfo",{}).get("milestoneDates",[])
                for m in mlist:
                    if m.get("type")=="ROUTE_LOCATION_MILESTONE" and m.get("key",{}).get("routeLocationType")=="ARRIVAL" and m.get("key",{}).get("locationMilestoneType")=="VEHICLE_ARRIVED":
                        eta=m.get("actualAchievementDateTime",{}).get("dateTime") or m.get("plannedAchievementDateTime",{}).get("dateTime"); break
                if eta:
                    eta=pd.to_datetime(eta,errors="coerce")
                    eta=None if pd.isna(eta) else eta.strftime("%d-%m-%Y")
                summary.append({"Shipment Number":sid,"Tracking Number":trk,"Container Number": c.get("containerNumber"),"Final ETA":eta,"Current Stage":c.get("routing",{}).get("milestoneInfo",{}).get("currentStage")})
                for rt in c.get("routing",{}).get("core",{}).get("routeLocations",[]):
                    x=pd.json_normalize(rt).to_dict("records")[0];x["Shipment ID"]=sid;x["Tracking Number"]=trk;routes.append(x)
                for ms in mlist:
                    x=pd.json_normalize(ms).to_dict("records")[0];x["Shipment ID"]=sid;x["Tracking Number"]=trk;milestones.append(x)
        for role,val in sh.get("parties",{}).items():
            vals=val if isinstance(val,list) else [val]
            for v in vals:
                if isinstance(v,dict):
                    x=pd.json_normalize(v).to_dict("records")[0];x["Role"]=role;x["Shipment ID"]=sid;x["Tracking Number"]=trk;parties.append(x)
    return {
      "Shipments":pd.json_normalize(details),
      "Containers":pd.DataFrame(containers),
      "Container Summary":pd.DataFrame(summary),
      "Routes":pd.DataFrame(routes),
      "Milestones":pd.DataFrame(milestones),
      "Parties":pd.DataFrame(parties),
      "Raw JSON":pd.DataFrame({"Shipment ID":[d.get("shipmentId") for d in details],"Tracking Number":[d.get("trackingNumber") for d in details],"Raw JSON":[json.dumps(d) for d in details]})
    }

def main():
    s=auth()
    ships=get_shipments(s)
    details=[]
    with ThreadPoolExecutor(max_workers=10) as ex:
        futs={ex.submit(fetch,s,x["shipmentId"]):x["shipmentId"] for x in ships}
        for f in as_completed(futs):
            d=f.result()
            if d: details.append(d)
    tables=build(details)
    with pd.ExcelWriter("Py_Shipments.xlsx",engine="openpyxl") as w:
        for n,df in tables.items(): df.to_excel(w,sheet_name=n,index=False)
    print("Done.")

if __name__=="__main__":
    main()
