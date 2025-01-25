import pandas as pd
import re

# file path
influencer_list_file_path = 'influencer_list_file_path'
label_fit_file = pd.ExcelFile('label_fit_file_path')
label_fit = {sheet_name: label_fit_file.parse(sheet_name) for sheet_name in label_fit_file.sheet_names}

pattern = u'[\\s\\d,.<>/?:;\'\"[\\]{}()\\|~!\t"@#$%^&*\\-_=+a-zA-Z，。\n《》、？：；“”‘’｛｝【】（）…￥！—┄－]+'

def calculate_weights(tags, brand):
    words = tags.split()
    weight = 0
    product_weighth = 0
    non_product_weight = 0
    count_total = len(words)
    count_product = label_fit[brand].loc[label_fit[brand]['product_or_image'] == 0, 'fit_sum'].values[0]
    count_non_product = label_fit[brand].loc[label_fit[brand]['product_or_image'] == 1, 'fit_sum'].values[0]
    count_total = count_product + count_non_product
    for word in words:
        if word in label_fit[brand]['label'].values:
            type = label_fit[brand].loc[label_fit[brand]['label'] == word, 'product_or_image'].values[0]
            fit = label_fit[brand].loc[label_fit[brand]['label'] == word, 'fit'].values[0]
            fit_total = label_fit[brand].loc[label_fit[brand]['label'] == word, 'fit_sum'].values[0]
            if type == 0:
                product_weighth += fit
            elif type == 1:
                non_product_weight += fit
            weight += fit
    return weight/count_total if count_total > 0 else 0, product_weighth/count_product if count_product > 0 else 0, non_product_weight/count_non_product if count_non_product > 0 else 0

with pd.ExcelWriter('influencer_fit_save_path') as writer:
    sheet_names = pd.ExcelFile(influencer_list_file_path).sheet_names

    for sheet_name in sheet_names:
        data = pd.read_excel(influencer_list_file_path, sheet_name)
        if "label" in data.columns:
            data['label'] = (
                data['label']
                .apply(lambda x: str(x))
                .apply(lambda x: re.sub(pattern, ' ', x))
            )
            data['total_fit'], data['product_fit'], data['image_fit'] = zip(*data.apply(lambda row: calculate_weights(row['label'], row['brand']), axis=1))
            data.to_excel(writer, sheet_name=sheet_name, index=False)