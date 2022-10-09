def createJson():
    import json
    import pandas as pd
    df = pd.read_excel('result.xlsx')
    # create json with indent   
    df.to_json('result.json', orient='records', indent=4)

createJson()