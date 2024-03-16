How to use Analyzer:
1. Cereate folder
2. create overview.xlsx
3. if you wish to exclude IPs like local ones -> create sheet named excludeIP, name first column IP and write them down
4. Download apache analyzer script
5. Run Apache Analyzer
6. 

only Sheet wich is getting changed is rqst, you can create new sheets and do something like i did and combine it with geo info


How to use geo Info:
1. Use apache analyser at least once
2. create an Sheet in overview.xlsx wich named "IPs"
3. copy all Ips From rqst and remove duplicates with excel
4. run geo_info.py
5. u can use VLOOKUP to move them in overview.xlsx

