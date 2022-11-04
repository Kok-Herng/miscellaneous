import pandas as pd 
import xml.etree.ElementTree as ET 
import itertools
  
def getAllElements(root): #contains ALL elements in file but no start and completion time for phases
    temp = []
    allElements = []

    for x in root.iter('RecipeElement'): 
        id = x.find('ID').text #find each element and get their text value
        desc = x.find('Description').text
        recipeElementType = x.find('RecipeElementType').text
    
        temp = [id, desc, recipeElementType]
        allElements.append(temp)
        
    allElements = pd.DataFrame(allElements, columns=['ID', 'Description', 'RecipeElementType']) #convert to dataframe
    return allElements

def getAllPhases(root): #contains ALL phases regarding of unit procedure
    allPhases = []

    for x in root.findall('.//ControlRecipe/RecipeElement/RecipeElement/RecipeElement/RecipeElement'): 
        id = x.find('ID').text
        desc = x.find('Description').text
        recipeElementType = x.find('RecipeElementType').text
        completionTime = x.find('CompletionTime').text
        startTime = x.find('StartTime').text

        temp = [id, desc, recipeElementType, completionTime, startTime]
        allPhases.append(temp)
    
    allPhases = pd.DataFrame(allPhases, columns=['ID', 'Description', 'RecipeElementType', 'CompletionTime', 'StartTime']) #convert to dataframe
    return allPhases

def calculatePhaseDuration(df):
    df['CompletionTime'] = pd.to_datetime(df['CompletionTime'], utc=True) #convert to datetime data type
    df['StartTime'] = pd.to_datetime(df['StartTime'], utc=True) #convert to datetime data type
    df['Duration'] = (df['CompletionTime'] - df['StartTime']) #duration between CompletionTime and StartTime

    finalDf = df.astype({'CompletionTime':'string','StartTime':'string'}) #convert CompletionTime and StartTime to string for Excel output later

    return finalDf

def pairwise(iterable): #to iterate the index by pairs, since phases are located between operations in our dataframe
    #########EXAMPLE##########
    ###OPERATION A###
    ###PHASE A1####### ------extracting this and put under OPERATION A
    ###PHASE A2####### ------extracting this and put under OPERATION A
    ###OPERATION B###
    ###PHASE B1####### ------extracting this and put under OPERATION B
    ###OPERATION C###
    a, b = itertools.tee(iterable)
    next(b, None)
    return zip(a, b) #eg: (0,1) (1,2) (2,3)

def opPhasesSplit(df): #split each operation and their phases
    operation = df.index[df['RecipeElementType'] == 'Operation'].tolist() #row index of all row containing the RecipeElementType 'operation'
    operation.append(df.tail(1).index.item()) #append the index of last record so the last operation is not missed out
    dfList = []
    
    for i, j in pairwise(operation):
        if j == operation[-1]: #if record is last row in dataframe
            dfList.append(df[i:j+1]) #include last row
        else: #if record is not last row in dataframe
            dfList.append(df[i:j]) #include up till the row before next operation
    return dfList

def totalTime(dfList): #find total time for each operation
    dfList2 = dfList

    for df in dfList2:
        df.loc['Total', 'Duration'] = df['Duration'].sum() #sum of duration of all phases for each operation
        df['ID'].iloc[-1] = 'Total' #change last row value to Total
        indexType = df[(df['RecipeElementType'] == 'UnitProcedure') | (df['RecipeElementType'] == 'Other') ].index #index of all UnitProcedure and Other
        df.drop(indexType , inplace=True) #remove all UnitProcedure and Other
    return dfList2

def totalTimeSummary(dfList): #summary of time taken for each operation in a dataframe
    summary = pd.DataFrame(columns=['ID', 'Description', 'CompletionTime', 'StartTime', 'Duration'])

    for total in dfList:
        total['Duration'].iloc[0] =  total.iloc[-1, -1] #copy total time from last row last column
        summary = summary.append(total.iloc[[0]], ignore_index=True) #add each operation row to dataframe
    return summary

def main():
    tree = ET.parse('dummyRecipe.xml') #1 procedure, 3 unit procedure, multiple operation and phase
    root = tree.getroot()

    allElements2 = getAllElements(root) #get all elements and put in single dataframe
    allPhases2 = getAllPhases(root) #get all phases and put in single dataframe
    phaseDuration = calculatePhaseDuration(allPhases2) #calculate duration for each phase
    finalDf = allElements2.merge(phaseDuration, how = 'left') #merge all elements and phases with duration together in single dataframe
    splittedList = opPhasesSplit(finalDf) #split phases according to their operation
    totalTimeList = totalTime(splittedList) #find total time for each operation
    totalTimeSummary2 = totalTimeSummary(totalTimeList) #combine time taken for each operation in a summary

    for df in totalTimeList:
        df['Duration'].iloc[0] = "" #remove total time from first row for each operation for clearer output

    totalTimeList.insert(0, totalTimeSummary2) #insert time summary for each operation at the front of list

    #for df in totalTimeList:
        
    with pd.ExcelWriter('operationTotalTime.xlsx', ) as writer: #export as Excel file
        for n, df in enumerate(totalTimeList):
            df['Duration'] = df['Duration'].astype(str) #convert Duration column to string for clearer Excel output
            if n == 0: #first sheet name is 'Summary'
                df.to_excel(writer, 'Summary', index = False)
            else: #second sheet onwards are operations
                df.to_excel(writer,'Operation%s' % n, index = False)

if __name__ == "__main__":
    main()