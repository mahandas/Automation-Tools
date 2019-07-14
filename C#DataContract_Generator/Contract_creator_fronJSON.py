import json 


var_json= """
 {
      "OrderID": "2454",
      "ClOrdID": "55112",
      "ExecID": "256899",
      "ExecTransType": "2",
      "ExecType": "5",
      "OrdStatus": "2",
      "LeavesQty": "2",
      "CumQty": "22",
      "Price": "100",
      "AvgPx": "20.2",
      "Text": "Some value"
   }
"""



vart = json.loads(str(var_json))




for k,v in vart.items():
  print("public string " + k + "  {get; set;}" )
