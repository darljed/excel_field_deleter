import sys, pandas as pd, json, os, time, asyncio, datetime





class Deleter:
    def __init__(self):
        self.pwd = os.path.dirname(os.path.realpath(__file__))
        self.debugMode = True
        self.output_filename = f'output-{datetime.datetime.now().strftime("%Y-%m-%dT%H-%M-%S")}'
        self.mainData = []
        self.tasks = []
        try:
            self.filename = sys.argv[1]
            if(len(sys.argv)>2):
                raise ValueError()
        except Exception as e:
            self.debug(f"Invalid or not declared file name. Please provide a valid file name of the source excel file as an argument.\nMake sure that the file name doesn't contain white spaces or try enclosing them in double quotes.")
            exit(1)

        asyncio.run(self.main())

    def debug(self,str):
        if(self.debugMode):
            print(str)
    
    def saveOutput(self,dataArr):
        with pd.ExcelWriter(os.path.join(self.pwd,f'{self.output_filename}.xlsx')) as writer:
            for data in dataArr:
                df = pd.DataFrame(data['data'])
                df.to_excel(writer,data['name'])

    async def processWorksheet(self,sheet):
        self.debug(f"{self.filename} : Parsing fields for worksheet {sheet}... This might take a few minutes to complete.")
        sheet_content = self.xl.parse(sheet)
        j = sheet_content.to_dict()
        self.debug(f"{self.filename} : removing unncessary fields...")
        # remove fields here
        newDict = {}
        for key in j:
            if key in self.fields_to_keep:
                newDict[key] = j[key]
        self.debug(f"{self.filename} : Cleaning up data...")
        newObj = {
            "name": sheet,
            "data": newDict
        }
        self.mainData.append(newObj)
        self.debug(f"{self.filename} : worksheet {sheet} has been cleaned-up.")

    async def main(self):
        
        # start the process
        startime = time.time()
        # capture list of fields that will be kept from a file 
        self.debug("Checking for fields_to_keep.txt ...")
        self.fields_to_keep = []
        if os.path.exists(os.path.join(self.pwd,"fields_to_keep.txt")):
            file = open('fields_to_keep.txt','r')
            self.fields_to_keep = file.read().split("\n")
            file.close()
            if len(self.fields_to_keep) == 0:
                self.debug("fields_to_keep.txt is empty. Aborting execution.")
            else:
                self.debug(f"Found {len(self.fields_to_keep)} fields to be kept.")
        else:
            self.debug("fields_to_keep.txt file doesn't exists. Aborting execution.")
            exit(1)
        # open the original excel file
        self.debug(f"Opening excel file {self.filename}")
        if(os.path.exists(os.path.join(self.pwd,self.filename))):
            self.xl = pd.ExcelFile(os.path.join(self.pwd,self.filename))
            self.debug(f"{self.filename} has been successfully loaded.")
            self.debug(f"{self.filename} : Reading worksheet names")
            # read the sheet names
            sheetnames = self.xl.sheet_names
            self.debug(f"{self.filename} : Found {len(sheetnames)} worksheet.")
            for sheet in sheetnames:
                # create a task
                self.tasks.append(asyncio.create_task(self.processWorksheet(sheet)))

        else:
            self.debug(f"The file {self.filename} does not exists. Aborting execution.")
            exit(1)

        await asyncio.gather(*self.tasks)
        self.debug("Almost there! Please wait while the new file is being generated... This will take some time for large datasets.")
        self.saveOutput(self.mainData)
        # with open('file.json',"w") as f:
        #     f.write(json.dumps(self.mainData))
        endtime = time.time()
        self.debug(f"Success! The clean up process has been completed within {endtime - startime} seconds.")
        self.debug(f"The new file {self.output_filename}.xlsx has been generated.")

    
        

if __name__ == "__main__":
    c = Deleter()
    