from cmath import nan
import pandas as pd
import os
import requests
import codecs
import re
import pandas as pd
import concurrent.futures



def read_thorlabs_cart(file, path=""):
    # Extract information about the file, look into current directory if empty
    if not path:
        path = os.path.abspath(os.getcwd())
    filepath = os.path.join(path, file)
    _, ext = os.path.splitext(filepath)

    # Load the file content as a dataframe
    if ext == '.xlsx' or ext == '.xls':
        df = pd.read_excel(filepath)
    elif ext == '.csv':
        df = pd.read_csv(filepath)
    else :
        df = pd.read_table(filepath)

    return df



def create_thorlabs_url(product):
    return "https://www.thorlabs.com/thorproduct.cfm?partnumber=" + str(product)



def get_thorlabs_product_weight(product, range=55, print_out=True, debug=False, file_out=False):
    # Get the web page source code
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36",
        "referer": "https://www.thorlabs.com"
    }
    page = requests.get(create_thorlabs_url(product), headers=headers)
    # print(page.status_code)
    data = page.content.decode('utf-8', 'replace')

    # Extract the relevant data
    idx = data.find('Poids total')
    substring = data[idx:idx+range]
    extracted = re.search(pattern="align=\"left\">(.*?)kg", string=substring)
    if extracted is None:
        weight = nan
    else:
        weight = float(extracted.group(1))

    if debug:
        print('\n substring = \n' + substring)
        print('\n extracted = \n' + extracted)

    if file_out:
        with codecs.open("test.txt", mode='w', encoding='utf-8') as f:
            f.write(page.content.decode('utf-8', 'replace'))

    if print_out:
        print(f"{product} weighs {weight} kg")

    return weight



def get_thorlabs_weights_from_cart(df):
    if 'Produit' and 'Product URL' in df.columns:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            results = executor.map(get_thorlabs_product_weight, df['Produit'])
        return list(results)
    else:
        return [None] * df.shape[0]



def append_weight_to_thorlabs_cart(df, weight):
    dfw = pd.DataFrame(weight)
    dfw.columns = ["Sub-Weight [kg]"]

    df = pd.concat([df, dfw], axis=1)
    if "Quantité" in df.columns:
        df["Weight [kg]"] = df["Quantité"] * df["Sub-Weight [kg]"]
    elif "Quantity" in df.columns:
        df["Weight [kg]"] = df["Quantity"] * df["Sub-Weight [kg]"]
    
    return df



def save_thorlabs_cart(df, path="", file="shoppingCartWeights.xls"):
    # Extract information about the file, look into current directory if empty
    if not path:
        path = os.path.abspath(os.getcwd())
    filepath = os.path.join(path, file)
    _, ext = os.path.splitext(filepath)

    # Save dataframe to file
    if ext == '.xlsx' or ext == '.xls':
        df.to_excel(filepath, index=False, float_format="%.2f")
    elif ext == '.csv':
        df.to_csv(filepath)
    else:
        df.to_csv()



if __name__ == '__main__':
    cart = read_thorlabs_cart("shoppingCart.xls", path="")
    weight = get_thorlabs_weights_from_cart(cart)
    cart = append_weight_to_thorlabs_cart(cart, weight)
    save_thorlabs_cart(cart, file="shoppingCartWithWeights.xls")
