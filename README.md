1. folder save C:\Windows\System32\cmd.exe
2. lineup and job_id_active ..these 2 file paths have to be saved in line 10 and 11
3. mattching is happenign on the abssis of composti_key ..so tha only 2 columns used are job_id / candidate_id  and comoposit_key of both
4. output goes in matchfiles forlder.. so delete any file in that folder before running python
5. pip install pandas execute
6. if any pip not installed then error will be thrown and install all those
7.logic goes as follows 
a) system goes via Job_id each ..and then matches the candidate_id.. 
b) output thrown is of lineup data
c) composit_key line up in column AG
d) 


Work to be done
1. add additional columns from job_id data 
job_id_composit_key	
Date	
company	HR	
HR_Phone	
Designation	
Age	
Education	
Gender

2. change code to consider a file which has both active and inactive but filter only by Active.. in column B in job_id

3. Date format has to be matching with sql format which is yyyy-mm-dd hh:mm:ss in job_id Column C and line up file column B (lineup format has to change to sql format permanently) dependign on your personal excel - 

4. match candidate unique but job_id multiple (and not company id).. so it is matching of canddiate_id with Job_id - this is an overkill - pause not to do now

5. 

