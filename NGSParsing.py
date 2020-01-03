import pandas as pd
import numpy as np

filename = "File1_input.xlsm"
df = pd.read_excel("File1_input.xlsm")
filename = filename.split("_")[0]


# first: sort by sample
def sort_by_sample() -> list:
    sample_col = df["Sample"]
    current_sample = sample_col[0]
    sorted_samples = []
    for sample in sample_col:
        if sample != current_sample:
            sorted_samples.append(df.loc[df["Sample"] == current_sample])
            current_sample = sample
    return sorted_samples


def get_types(sample: pd.Series) -> tuple:
    pathogen_column = sample["Pathogen_reference"]
    subtypes = []
    genes = []
    for pathogen_ref in pathogen_column:
        split_path_ref = pathogen_ref.split("|")
        if len(split_path_ref) >= 4:
            subtype = split_path_ref[3]
            gene_name = split_path_ref[1]
            subtypes.append(subtype)
            genes.append(gene_name)
    return subtypes, genes


def add_labels_to_types(types: list) -> list:
    cols = []
    for col in df:
        cols.append(col)
    labelled_types = []
    for col in cols[2:]:
        for _type in types:
            _type = _type + "_" + col
            labelled_types.append(_type)
    labelled_types.sort()
    return labelled_types


def add_columns_to_df(data: pd.Series) -> pd.Series:
    pathogen_col = data["Pathogen_reference"]
    pathogens = []
    types = []
    subtypes = []
    genes = []
    for entry in pathogen_col:
        entry = entry.split("|")
        if len(entry) >= 4:
            types.append(entry[2])
            subtypes.append(entry[3])
            pathogens.append(entry[0])
            genes.append(entry[1])
        else:
            pathogens.append("not_applicable")
            types.append("not_applicable")
            subtypes.append("not_applicable")
            genes.append("not_applicable")
    data.insert(1, "Pathogen", pathogens)
    data.insert(2, "Type", types)
    data.insert(3, "Subtype", subtypes)
    data.insert(4, "Gene", genes)
    del data["Pathogen_reference"]
    return data


def make_new_df(column_names: list, df_array: np.array) -> dict:
    sorted_df = {}
    for i in range(len(column_names)):
        sorted_df[column_names[i]] = []
    sorted_df["Sample"].append(df_array[0][0])
    sorted_df["Pathogen"].append(df_array[0][1])
    final_subtype = ""
    for sublist in df_array:
        if sublist[1] == "not_applicable":
            sorted_df["not_applicable_breadth_coverage"].append(0)
            sorted_df["not_applicable_RPK"].append(0)
            sorted_df["not_applicable_RPKM"].append(0)
            sorted_df["not_applicable_read_counts"].append(0)
        elif sublist[3] != "not_applicable":
            subtype = sublist[3]
            bc_key = subtype + "_breadth_coverage"
            sorted_df[bc_key].append(sublist[-1])
            RPKM_key = subtype + "_RPKM"
            sorted_df[RPKM_key].append(sublist[-2])
            RPK_key = subtype + "_RPK"
            sorted_df[RPK_key].append(sublist[-3])
            rc_key = subtype + "_read_counts"
            sorted_df[rc_key].append(sublist[-4])
            final_subtype = final_subtype + subtype
        else:
            gene_name = sublist[4]
            bc_key = gene_name + "_breadth_coverage"
            sorted_df[bc_key].append(sublist[-1])
            RPKM_key = gene_name + "_RPKM"
            sorted_df[RPKM_key].append(sublist[-2])
            RPK_key = gene_name + "_RPK"
            sorted_df[RPK_key].append(sublist[-3])
            rc_key = gene_name + "_read_counts"
            sorted_df[rc_key].append(sublist[-4])
        if sublist[2] != "not_applicable":
            if sublist[2] not in sorted_df["Type"]:
                sorted_df["Type"].append(sublist[2])

    sorted_df["Subtype"].append(final_subtype)
    for key in sorted_df:
        if len(sorted_df[key]) == 0:
            sorted_df[key].append("")

    return sorted_df


if __name__ == "__main__":
    sorted_samples = sort_by_sample()
    all_subtypes = []
    all_genes = []
    # print(sort_by_type(sorted_samples[0]))
    for i in range(len(sorted_samples)):
        types_from_df = get_types(sorted_samples[i])
        all_subtypes = all_subtypes + types_from_df[0]
        all_genes = all_genes + types_from_df[1]

    # removing duplicate types and subtypes
    all_subtypes = list(dict.fromkeys(all_subtypes))
    all_genes = list(dict.fromkeys(all_genes))

    all_types = all_genes + all_subtypes
    col_headers = add_labels_to_types(all_types)
    default_col_headers = ["Sample", "Pathogen", "Type", "Subtype"]
    col_headers = default_col_headers + col_headers

    new_dfs = []
    for i in range(len(sorted_samples)):
        sorted_samples[i] = add_columns_to_df(sorted_samples[i])
        sorted_samples[i] = sorted_samples[i].to_numpy()
        new_dfs.append(make_new_df(col_headers, sorted_samples[i]))

    # setting up the final dictionary; initializing it to have the correct keys
    final_dict = {}
    for header in col_headers:
        final_dict[header] = []

    # combining the values of all the smaller dictionaries into one big dict
    for _dict in new_dfs:
        for key in _dict:
            final_dict[key].extend(_dict[key])

    final_series = pd.DataFrame(final_dict)
    outfile = open(filename + "_output" + ".xlsx", "wb+")
    final_series.to_excel(outfile)
    outfile.close()

