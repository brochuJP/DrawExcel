import pandas as pd
import re
import win32com.client
from graphviz import Digraph

def LoadExcelStructure(fileFolder,fileName):
    """
    Return a dataframe containing information about your Excel file VB structure
    fileFolder: Your Excel file folder
    fileName: Your Excel file name including the extension
    """
    fileFolder=fileFolder + "/"
    xl = win32com.client.Dispatch('Excel.Application')
    wb = xl.Workbooks.Open(fileFolder + fileName)
    xl.Visible = 1
    df_ModInfo=pd.DataFrame()
    listInfo=[]
    # Go over every VBA compoents
    print("Reading the Excel structure")
    for VBComp in wb.VBProject.VBComponents:
        for lineCodeMod in range(1,VBComp.CodeModule.countOfLines):
            VBComponent=VBComp.name
            VBComponentClean=re.sub(r'[\W_]+','',VBComponent)
            ProcName=VBComp.CodeModule.ProcOfLine(lineCodeMod)
            ProcNameClean=re.sub(r'[\W_]+','',ProcName[0])
            ProcLineNumber=lineCodeMod-VBComp.CodeModule.ProcStartLine(ProcName[0],ProcName[1])
            ProcLineNumberFromBody=lineCodeMod-VBComp.CodeModule.ProcBodyLine(ProcName[0],ProcName[1])
            LineOfCode=VBComp.CodeModule.Lines(lineCodeMod,1)
            listInfo.append([VBComponent,VBComponentClean,ProcName[0],ProcNameClean,ProcName[1],ProcLineNumber,ProcLineNumberFromBody,LineOfCode])
            VBComp.CodeModule
            df_ModInfo
    df_ModInfo=pd.DataFrame(listInfo,columns=['VBComponent','VBComponentClean','ProcName','ProcNameClean','ProcKind','ProcLineNumber','ProcLineNumberFromBody','LineOfCode'])

    df_ModInfo['FoncOnLine'] = FoncOnLine(df_ModInfo['LineOfCode'],df_ModInfo['ProcName'],df_ModInfo['ProcLineNumberFromBody'])
    wb.Close(False)
    return df_ModInfo

def FoncOnLine(series,ProcName,ProcLineNumberFromBody):
    """
    Check if a function or a sub is being called in a line of code
    series: List of all lines of code
    ProcName : All the proceduers names
    ProcLineNumberFromBody: The line number of the function of sub
    """
    names=set(ProcName)

    # FINDS ANY NAME
    matches_DefaultLink= pd.DataFrame(data=None)
    for name in names:
        matches_DefaultLink[name]=series.str.contains(name)
        booMsk=list(matches_DefaultLink[(name)])
        matches_DefaultLink[name].loc[booMsk] = pd.Series(ProcName) + '->' + pd.Series([name]*len(ProcName))
    
    # FINDS FUNCTION DEFINITIONS
    matches_def = pd.DataFrame(data=None)
    for name in names:
        pattern = rf'{name}\(\)'
        matches_def[name] = series.str.contains(pattern, regex=True)

    # FINDS STRINGS
    matches_str = pd.DataFrame(data=None)
    for name in names:
        pattern = rf'\".*{name}.*\"'
        matches_str[name] = series.str.contains(pattern, regex=True)

    # FINDS IF IS HEADER
    matches_head = pd.DataFrame(columns=matches_str.columns,index=matches_str.index)
    matches_head.iloc[:,:]=False
    matches_head[ProcLineNumberFromBody==0]=True
 
    # Check if the we preserve the link
    matches_to_preserve=(matches_DefaultLink!=False) & ~matches_def & ~matches_str & ~matches_head
    matches=matches_DefaultLink[matches_to_preserve]

    # Put the result in camma separed unique values
    matches_lists=pd.Series(matches.fillna('').values.tolist())
    matches_lists=[",".join([val for val in matches_list if val!='']) for matches_list in matches_lists]

    return matches_lists

def drawA_Graph(dfToDraw,fileFolder,fileName):
    """
    create multiple graphs depending of the granularity of the vb Objects
    dfToDraw: Raw DataFrame of your Excel file structure
    fileFolder: Your Excel file folder
    fileName: Your Excel file name including the extension
    """
    dot = Digraph(comment=fileFolder + fileName,format='svg',engine='dot')
    dot.graph_attr.update(size="8.3,11.7!",margin="0.5,0.5",spline='true',sep=".01",ranksep="1",rankdir='RL',
                            layout="fdp",overlap='false', center='true',splines='true')
    # Declare the dot in there cluster
    # If we have more then when cluster then we draw the cluster if not then we do not draw
    df_SubInfo=dfToDraw[['VBComponent','VBComponentClean','ProcName','ProcNameClean']].drop_duplicates()
    # CREATE SUBGRAPH AND ASSIGN NODES IF APPLICABLE
    booSubMode = True if len(df_SubInfo['VBComponentClean'].unique())>1 else False
    if booSubMode:
        for curVBComp in dfToDraw['VBComponentClean'].drop_duplicates():
            df_Sub_Proc=df_SubInfo.loc[curVBComp==dfToDraw['VBComponent'],'ProcNameClean']
            with dot.subgraph(name='cluster_' + curVBComp) as curVBCompObj:
                curVBCompObj.attr(label=str(curVBComp),href=fileFolder + '/' + fileName.split(".")[0] + '/' + str(curVBComp) + ".gv.svg")
                [curVBCompObj.node(x) for x in df_Sub_Proc.to_list()]

    [dot.node(CleanProcName, RealProcName) for CleanProcName, RealProcName in zip(dfToDraw['ProcNameClean'], dfToDraw['ProcNameClean'])]

    # CREATE ARROWS CONNECTIONS
    for string in dfToDraw['FoncOnLine']:
        if string!='':
            for strLink in string.split(','):
                lstOrgDest=list(strLink.split('->'))
                lstOrgDestClean=[NameCleaner(strName) for strName in lstOrgDest ]
                dot.edge(lstOrgDestClean[0],lstOrgDestClean[1],constraint='false')

    if booSubMode==True:
        dot.render(fileFolder + '/' + fileName.split(".")[0] + "/_MAIN" + ".gv")
        # Run each VBComponent
        AllVBComp=dfToDraw['VBComponentClean'].drop_duplicates()
        for curVBComp in AllVBComp:
            drawA_Graph(dfToDraw[dfToDraw['VBComponentClean']==curVBComp],fileFolder,fileName)
    else:
        dot.render(fileFolder + '/' + fileName.split(".")[0] + '/' + df_SubInfo['VBComponentClean'].unique()[0] + ".gv")

def NameCleaner(strName):
    return re.sub(r'[\W_]+','',strName)

def DrawExcel(fileFolder,fileName):
    """
    Scan the VBA code in a MS Excel file and generate a visual diagram of it's structure.
    fileFolder: Your Excel file folder
    fileName: Your Excel file name including the extension
    """
    print("Opening the Excel file")
    df_XlStuct=LoadExcelStructure(fileFolder,fileName)
    print("Starting graph generations")
    drawA_Graph(df_XlStuct,fileFolder,fileName)
    print("Done")
def main():
    fileFolder=r"My Excel file folder Path"
    fileName = r"My Excel file name including extension"
    DrawExcel(fileName,fileFolder)
    print("Main representation availeble here: \n" + crlf + fileFolder + "/" + fileName + "/_Main.svg" ) 
if __name__== "__main__":
    main()