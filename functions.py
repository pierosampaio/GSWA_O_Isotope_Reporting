
import pandas as pd
import numpy as np
from scipy.stats import norm
import re

import os
import sys
from pathlib import Path


from scipy.optimize import brentq

from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment



def filter_by_regex(strings, pattern):
    """
    Returns a list of strings from `strings` that match the given regex `pattern`.
    
    Args:
      strings (list): List of strings to search.
      pattern (str): Regular expression pattern.
    
    Returns:
      list: Strings that match the pattern.
    """
    if not isinstance(strings, list) or not all(isinstance(s, str) for s in strings):
        raise ValueError("Input must be a list of strings.")
    if not isinstance(pattern, str):
        raise ValueError("Pattern must be a string.")
    
    try:
        regex = re.compile(pattern)
    except re.error as e:
        raise ValueError(f"Invalid regular expression: {e}")
    
    matches = [s for s in strings if regex.search(s)]
    
    if len(matches) == 1: return matches[0]
    elif not matches: return None
    else: print(matches)

def find_column(df, candidates):
    normalize = lambda s: re.sub(r'[\s_\-]+','',s).lower()
    normalized_columns = {normalize(c): c for c in df.columns}
    for name in candidates:
        match = normalized_columns.get(normalize(name))
        if match:
            return match
    return None

#def find_column(df, candidates):
#    matches = difflib.get_close_matches(
#        " ".join(candidates), df.columns.to_list(), n = 1, cutoff = 0.4
#    )
#    return matches[0] if matches else None


def UPb_xls_process(UPb_path,UPb_file):

    header = [
      "GroupID","SpotNo","GrainSpot","238U_ppm","232Th_ppm",
      "Th232U238","f204_pct","U238Pb206","U238Pb206_1sig",
      "Pb207Pb206","Pb207Pb206_1sig","4corr_U238Pb206",
      "4corr_U238Pb206_1sig","4corr_Pb207Pb206","4corr_Pb207Pb206_1sig",
      "4corr_86_date","4corr_86_date_1sig","4corr_76_date",
      "4corr_76_date_1sig","Discordance_pct"
    ]

    Sample_name = UPb_file.split(".")[0]
    Sample_name = Sample_name.removesuffix("-combined")
    
    xl = pd.ExcelFile(os.path.join(UPb_path,UPb_file))
    sheet_names = xl.sheet_names
    target_sheets = ["excel_table","data_table"]
    found_sheet = next((s for s in target_sheets if s in sheet_names), None)
    skiprows = {
      "excel_table":1,
      "data_table":4
    }
    
    df = pd.read_excel(os.path.join(UPb_path,UPb_file), sheet_name = found_sheet, skiprows = skiprows[found_sheet])
    
    ## Sanitise dataframe
    df = df.iloc[:,:20]
    df = df.dropna(axis=0)
    
    
    df = df.rename(
    dict(zip(df.columns,header)),
    axis=1
    )
    
    
    
    if found_sheet == "data_table":
        df["Spot"] = [str(v).replace(".","-") for v in df["GrainSpot"].values]
    elif found_sheet == "excel_table":
        df["Spot"] = df["GrainSpot"].str.split("-", expand=True)[1].str.replace(".","-")
    
    
    df["UPB_ANALYSIS_ID"] = df.GrainSpot
    
    
    df["Sample"] = Sample_name
    
    df["SampleSpot"] = (df["Sample"] + "-" + df["Spot"])
    
    df["SampleSpot"] = df["SampleSpot"].str.replace(".","-")
    
    df["SpotNo"] = df["SpotNo"].astype("int").astype("string")
    
    
    
    return(df)


def UPb_txt_process(UPb_path,UPb_file):
  
    header = [
    "Sample","GroupID","SpotNo","UPB_ANALYSIS_ID","238U_ppm","232Th_ppm",
    "Th232U238","f204_pct","U238Pb206","U238Pb206_1sig",
    "Pb207Pb206","Pb207Pb206_1sig","4corr_U238Pb206",
    "4corr_U238Pb206_1sig","4corr_Pb207Pb206","4corr_Pb207Pb206_1sig",
    "4corr_86_date","4corr_86_date_1sig","4corr_76_date",
    "4corr_76_date_1sig","Discordance_pct","SampleSpot","GrainSpot"
    ]

    
    
    #### Create regular expression to find the correct Grain spot (case sensitive) and
    #### Group ID columns
    
    
    df = pd.read_table(
      os.path.join(
          UPb_path,UPb_file
      )
    )

    group_col_cands = ["Group ID", "Grp no", "Group No", "Group no", "Grp ID", "Grp No"]
    grainspot_col_cands = ["Grain Spot", "grain spot", "GrainSpot", "Grainspot", "Grain spot"]
    spotno_col_cands = ["Spotno", "spotno", "Spot no"]

    
    group_col = find_column(df, group_col_cands)
    grainspot_col = find_column(df, grainspot_col_cands)
    spotno_col = find_column(df, spotno_col_cands)

    if len(df.columns) > 20:
    
        df["Geochronid"] = df["Geochronid"].astype("string")

        if df[grainspot_col].dtype == "float64":
            df["Grain spot2"] = df[grainspot_col].astype("str").str.replace(".","-")
        else:
            df["Grain spot2"] = df[grainspot_col].str.split("-",expand=True)[1].str.replace(".","-")
            
        if df["Geochronid"].str.startswith("GSWA").any():
            df["SampleSpot"] = (
              df["Geochronid"].str.split("_", expand = True)[1].str.split(".", expand = True)[0] +
              "-" +
              df["Grain spot2"]
            ).astype("str")
        else: 
            df["SampleSpot"] = (
              df["Geochronid"].str.split(".", expand = True)[0] + "-" +
              df["Grain spot2"]
            ).astype("str")
        
        ## SANITISE AND RENAME COLUMNS
        
        df = df.rename({grainspot_col:"UPB_ANALYSIS_ID",
                        group_col: "Group ID",
                        spotno_col: "Spot no"}, axis = 1)
        
        df = df[[
          "Geochronid", "Group ID", "Spot no", "UPB_ANALYSIS_ID", "238U(ppm)",
          "232Th(ppm)", "232Th_238U", "f(%)", "238U_206Pb", "238U_206Pb_er",
          "207Pb_206Pb", "207Pb_206Pb_er", "238U_206Pb*", "238U_206Pb*_er",
          "207Pb*_206*Pb", "207Pb*_206*Pb_er", "238U_206Pb*_age",
          "238U_206Pb*_age_er", "207Pb*_206Pb*_age", "207Pb*_206Pb*_age_er",
          "Disc(%)","SampleSpot"
        ]]
    
        
        df = df.rename(
          dict(zip(df.columns,header)),
          axis = 1 
        )

    else:
        print(f"Check {UPb_file} manually")
        return None
    
    return df
    
def UPb_file_join(UPb_path):

    UPb_dfs = []
    for f in os.listdir(UPb_path):
        if os.path.splitext(f)[-1] == ".xls": 
            UPb_dfs.append(UPb_xls_process(
              UPb_path,
              f
            ))
        elif os.path.splitext(f)[-1] == ".txt":
            UPb_dfs.append(UPb_txt_process(
            UPb_path,
            f
            ))

        else:
            pass
    assert len(UPb_dfs) > 0, "No files to concatenate"
    return pd.concat(UPb_dfs)


def Oxygen_processing(path,file):

    df = pd.read_excel(
        os.path.join(
            path,file
        ),
        sheet_name="3-CorrectedData",
        skiprows=9
    )

    df = df[[
        "Analysis #","18O/16O","± rel (%)",
        "16O1H/16O","± rel (%).1","d18O.1","± per mil.2",
        "OH/O","± rel (%).2"
    ]]

    df = df.dropna(
        how = "any",
        axis = 0
    )

    new_cols = [
        "Unique_O_ID", "18O/16O", "18O/16O_± rel (%)", "16O1H/16O","16O1H/16O_± rel (%)","d18O", "± per mil", "OH/O", "OH/O_± rel (%)"
    ]

    df = df.rename(dict(zip(df.columns,new_cols)), axis = 1)

    def get_samplespot_name(df):

        longest_str = lambda l: max(l, key=len)
        sample_spot = []
        sample = []
        for i, row in df.iterrows():
            
            longname = row["Unique_O_ID"]
            
            # get spot names    
            if "@" in longname:
                spotname = longname.split("@")[-1]
            else:
                spotname = "-".join(longname.split("-")[-2:])

            # get sample names

            first_guess = longest_str(longname.split("@")[0].split("-"))
            if "_" in first_guess:
                guess = first_guess.split("_")[1]
            else:
                guess = first_guess
            
            sample_spot.append(guess + "-" + spotname)
            sample.append(guess)
        return (sample_spot,sample)
    
    SampleSpot = get_samplespot_name(df)[0]
    Sample = get_samplespot_name(df)[1]

    df["SampleSpot"] = SampleSpot
    df["Sample"] = Sample
    df["file"] = file

    return df

def O_file_join(path):

    O_dfs = []
    print("Found Oxygen files:\n")
    for f in os.listdir(path):
        print(f)
        O_dfs.append(
            Oxygen_processing(
                path,f
            )
        )
    assert len(O_dfs) > 0, "No files to concatenate"
    return pd.concat(O_dfs)


def merge_dataset(UPb_dataset, O_dataset, joining_key):

    UPb_dataset[joining_key] = UPb_dataset[joining_key].astype("str")
    O_dataset[joining_key] = O_dataset[joining_key].astype("str")

    UPb_samples = {m.group() for s in UPb_dataset["Sample"].unique() if \
                   (m := re.search(r"\d{6}", s))}
    O_samples = {s for s in O_dataset["Sample"].unique() if s[0].isnumeric()} # Only append GSWA sample codes

    Problem = UPb_samples.difference(O_samples).union(O_samples.difference(UPb_samples))

    UPb_problem = Problem.intersection(UPb_samples)
    O_problem = Problem.intersection(O_samples)

    df_merged = pd.merge(
        O_dataset,UPb_dataset,on=joining_key,how="left"
    )

    if len(UPb_problem) > 0:
        print("\nSamples with U-Pb but no Oxygen data:")
        print("\n".join(UPb_problem))
    
    elif len(O_problem) > 0:
        print("\nSamples with Oxygen data but no U-Pb:")
        print("\n".join(O_problem))

    else:
        print("All samples matched!")
    
    return df_merged

def calculate_mswd(vals, errs, w_mean):
    n = len(vals)
    if n <=1: return 0
    weights = 1./ (errs**2)
    mswd = np.sum(weights * (vals - w_mean)**2)/(n-1)
    return mswd

def weighted_mean(values, uncertainties, method = "internal_sigma"):
    """
    Performs a weighted mean calculation with inverse variance weighting. Optionally
    performs outlier rejection.

    Parameters:
    - values: data points
    - uncertainties: 1-sigma uncertainties (SE or SD)
    - method: "chauvenet" (as in IsoplotR), "internal_sigma" (GSWA
    age calculations) or None (does not exclude outliers)
    """

    vals = np.array(values, dtype=float)
    errs = np.array(uncertainties, dtype=float)

    iteration = 0
    while True:
        n = len(vals)
        if n <= 1: break

        weights = 1.0/(errs**2)

        w_mean = np.dot(vals,weights)/np.sum(weights)

        # Calculate deviations
        deviations = np.abs(vals - w_mean)

        if method.lower() == "chauvenet":
        # spread-based: uses Z-score (deviation/population_std)
            std_dev = np.std(vals)
            if std_dev == 0: break
            z_scores = deviations/std_dev
            probs = n * (2 * (1 - norm.cdf(z_scores)))

            if np.min(probs) < 0.05:
                reject_idx = np.argmin(probs)
                do_reject = True
            else:
                do_reject = False

        elif method.lower() == "internal_sigma":
            # Uncertainty-based: deviation/individual error
            # Checks if the mean is > X sigma away from the point's own error
            sigma_ratios = deviations/errs
            max_ratio_idx = np.argmax(sigma_ratios)

            if sigma_ratios[max_ratio_idx] > 2.5:
                reject_idx = max_ratio_idx
                do_reject = True
            else:
                do_reject = False

        if do_reject:
            vals = np.delete(vals, reject_idx)
            errs = np.delete(errs, reject_idx)
            iteration += 1
        else:
            break

    final_w_mean = np.sum((1.0/(errs**2)) * vals) / np.sum(1.0/(errs**2))
    internal_err = np.sqrt(1.0/np.sum(1.0/(errs**2)))
    mswd = calculate_mswd(vals,errs,final_w_mean)
    external_err = internal_err * np.sqrt(mswd)
    ci = 1.96 * internal_err


    return {
        "mean": final_w_mean,
        "uncertainty": internal_err,
        "external_error": np.round(external_err,1),
        "ci": np.round(ci,1),
        "mswd": mswd,
        "data": vals,
        "n":len(vals),
        "n_rejected": iteration
    }
    


def calc_group_stats(df: pd.DataFrame, grouping_var: str, variable: str, variable_unc: str, prepend: str) -> pd.DataFrame: 
    """
    Takes a grouped dataframe and calculates group level aggregate statistics for a given
    variable, which should be the name of the column of interest and its associated uncertainty
    (variable_unc)
    """

    
    
    dfg = df.groupby(grouping_var)
    
    median = dfg[variable].median().round(2)
    median_1sigma = ((dfg[variable_unc].median()**2 + dfg[variable].std()**2).values)**0.5 # Following Yong Jun's method
    average = dfg[variable].mean().round(2)
    std = dfg[variable].std().round(2)
    maxx = dfg[variable].max()
    min = dfg[variable].min()

    sample_weighted_mean = (
            dfg
            .apply(lambda g: weighted_mean(
                g[variable],
                g[variable_unc]
            )
        ))
    
    sample_weighted_mean = pd.DataFrame.from_dict(sample_weighted_mean.to_dict(), orient="index")
    wm = sample_weighted_mean["mean"].round(2)
    uncertainty = sample_weighted_mean["uncertainty"].round(2)
    mswd = sample_weighted_mean["mswd"].round(2)
    n = sample_weighted_mean["n"]
    ci = np.full(len(dfg),"95%")

    data = np.array((
        median,median_1sigma,average,std,wm,uncertainty,n,mswd,ci
    ))

    headers = [
        "MDN","1SIG","AVE","1SD","WM","UNCER","PRIMARY_N","MSWD","CONF"
    ]
    
    index = sample_weighted_mean.index
    columns = [prepend + "_" + lab for lab in headers]

    return pd.DataFrame(data=data.T,columns=columns,index=index)


def create_aggregate_df(df: pd.DataFrame, grouping_var: str, vals: list, uncertainties: list, prepends: list, by_list = None) -> pd.DataFrame:

    df.insert(len(df.columns),"exclude",np.nan)
  
  
  
    if by_list is not None:
        with open(by_list, "r") as f:
            vals_to_exclude = [line.strip() for line in f if line.strip() and not line.startswith("#")]
  
  
    else:
        print("Select spots to exclude:")
        vals_to_exclude = []

        for val in df[grouping_var].unique():
            print(f"\n\nSpots in sample {val}:\n")

            available_spots = df.loc[df[grouping_var] == val, "SampleSpot"].tolist()
            print("\n".join(available_spots))
            
            print("Select spot number to exclude from calculation. Type just the spot number, without the sample.\nType 'X' to finish for this sample. Type 'Exit' to quit the function.")
            s = ""
            while s != "X":
                s = input("Spot number:")
                if s.upper() == "X": break
                if s.title() == "Exit": return None
                if (val.split(".")[0] + "-" + s) not in map(str,available_spots):
                    print("Invalid spot!")
                    continue
                vals_to_exclude.append(val.split(".")[0] + "-" + s)
            corr = input("Confirm selection? Press 'Y' to confirm or 'N' to select spots to remove from list.")
            if corr.upper() == "N":
                print("Select spot undo, press X to finish:")
                s = ""
                while s.upper() != "X":
                    s = input("Spot number:")
                    if s.upper() == "X": break
                    if s.title() == "Exit": return None
                    if (val.split(".")[0] + "-" + s) not in vals_to_exclude:
                        print("Valid not in exclude list.")
                        continue
                    vals_to_exclude.remove(val.split(".")[0] + "-" + s)
            
    print("Excluded spots:\n","\n".join(vals_to_exclude))
          

    df.loc[df["Unique_O_ID"].isin(vals_to_exclude),"exclude"] = 1

    df = df.loc[df["exclude"].isna()]

    dataframes = []
    for v,u,p in zip(vals, uncertainties, prepends):
        dataframes.append(calc_group_stats(df,grouping_var,v,u,p))

    master_df = pd.concat(dataframes, axis = 1)
  
    try:
        ages = df.loc[df.GroupID == "I"].groupby("Sample")["Age_Hf_calculation"].mean()
        sigs = df.loc[df.GroupID == "I"].groupby("Sample")["Age_Hf_calculation_unc"].mean()

        master_df["Age"] = master_df.index.map(ages)
        master_df["Age_1sig"] = master_df.index.map(sigs)

    except:
        pass
  
    return master_df

### U-Pb constants for calculating weighted mean ages

lbd235 = 9.8485e-10
lbd238 = 1.55125e-10
U238_U235 = 137.818

def ratio_function(t):
    return (1/U238_U235) * \
           (np.exp(lbd235*t) - 1) / \
           (np.exp(lbd238*t) - 1)

def age_from_ratio(ratio):
    def f(t):
        return ratio_function(t) - ratio
    return brentq(f, 1e6, 4.5e9)

def pb207_pb206_age_with_uncertainty(ratio, sigma_ratio):
    
    # Solve for age
    t = age_from_ratio(ratio)
    
    # Compute derivative dR/dt
    A = np.exp(lbd235*t)
    B = np.exp(lbd238*t)
    
    dRdt = (1/U238_U235) * (
        (lbd235*A*(B-1) - lbd238*B*(A-1)) /
        ((B-1)**2)
    )
    
    # Propagate uncertainty
    sigma_t = sigma_ratio / abs(dRdt)
    
    t *= 1e-6
    sigma_t *= 1e-6

	
    return t, sigma_t



