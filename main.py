import pandas as pd
import re
from datetime import datetime

file_path = r"C:\temp\Shortage Report.xls"
year = datetime.now().year


def map_sale_prices():
    global file_path

    df = pd.read_excel(file_path, sheet_name="Shortages2", header=None)
    df.columns = [
        "Kit",
        "Description",
        "Qty",
        "So Date",
        "Del Date",
        "SO Nbr",
        "Cust",
        "Cust Name",
        "Sell $",
    ]
    df.dropna(subset=["Cust"], inplace=True)
    df["SO Nbr"] = df["SO Nbr"].apply(
        lambda x: str(float(x)).rstrip(".0") if x not in ["", "End of"] else x
    )
    df["Unit Sell $"] = df["Sell $"] / df["Qty"]

    sell_prices_by_kit_so_nbr = {}
    for _, row in df.iterrows():
        key = f"{row['Kit']}_{row['SO Nbr']}"
        # print(key)
        if key not in sell_prices_by_kit_so_nbr:
            sell_prices_by_kit_so_nbr[key] = row["Unit Sell $"]

    return sell_prices_by_kit_so_nbr


mapped_prices = map_sale_prices()


def shortages1():
    global file_path

    def get_price(row):
        global mapped_prices
        key = f"{row['Kit']}_{row['SO Nbr']}"
        # print(key)
        return mapped_prices.get(key, 0)

    df = pd.read_excel(file_path, sheet_name="Shortages1", header=None)
    df.columns = [
        "SO Nbr",
        "Cust",
        "Cust Name",
        "Kit",
        "Description",
        "Qty",
        "Del Date",
        "Cust PO",
        "So Date",
    ]

    df.fillna("", inplace=True)

    df["SO Nbr"] = df["SO Nbr"].apply(
        lambda x: str(float(x)).rstrip(".0") if x not in ["", "End of"] else x
    )
    df["Kit"] = df["Kit"].astype(str)
    df["Cust PO"] = df["Cust PO"].astype(str)

    df["So Date"] = df["So Date"].astype(str).str[:10]
    df["So Date"] = df["So Date"].apply(lambda x: "" if x == "NaT" else x)
    df["Del Date"] = df["Del Date"].astype(str).str[:10]
    df["Del Date"] = df["Del Date"].apply(lambda x: "" if x == "NaT" else x)

    df["Unit Sell $"] = df.apply(get_price, axis=1)
    df["Sell $"] = df.apply(
        lambda row: row["Unit Sell $"] * row["Qty"] if row["Qty"] != "" else "",
        axis=1,
    )

    df["Unit Sell $"] = df["Unit Sell $"].apply(lambda x: "" if x == 0 else x)

    df.to_excel(
        f"Shortages1 (By SO) RunTime-{datetime.now(): %Y-%m-%d %H%M%S}.xlsx",
        index=False,
    )

    # df.dropna(subset=["SO Nbr"], inplace=True)


def shortages2():
    global file_path

    back_order_regex = re.compile(r"-12-31", re.IGNORECASE)
    special_case_1 = re.compile(r"-01-01", re.IGNORECASE)
    special_case_2 = re.compile(r"-12-12", re.IGNORECASE)
    special_case_3 = re.compile(r"-01-31", re.IGNORECASE)
    special_case_4 = re.compile(r"-12-01", re.IGNORECASE)

    def reason_from_date(date):
        special_code = int(date[:4]) > year

        if back_order_regex.search(date) and special_code:
            return "[12/31] This order will ship as soon as it is released..."
        elif special_case_1.search(date) and special_code:
            return "[01/01] This customer is required to pay in advance"
        elif special_case_2.search(date) and special_code:
            return "[12/12] The order is on hold and it is not clear when it will ship"
        elif special_case_3.search(date) and special_code:
            return (
                "[01/31] The rest of the order shipped but this line item is in dispute"
            )
        elif special_case_4.search(date) and special_code:
            return "[12/01] This customer is on credit hold"
        else:
            return f"This customer requested a specific delivery date: {date.replace('-', '/')}"

    df = pd.read_excel(file_path, sheet_name="Shortages2", header=None)
    df.columns = [
        "Kit",
        "Description",
        "Qty",
        "So Date",
        "Del Date",
        "SO Nbr",
        "Cust",
        "Cust Name",
        "Sell $",
    ]

    df.dropna(subset=["Cust"], inplace=True)

    df["Kit"] = df["Kit"].astype(str)
    df["So Date"] = df["So Date"].astype(str).str[:10]
    df["Del Date"] = df["Del Date"].astype(str).str[:10]

    df["Unit Sell $"] = df["Sell $"] / df["Qty"]

    df["Reason"] = df["Del Date"].apply(reason_from_date)

    grouped_by_kit = df.groupby("Kit").agg(
        {
            "Qty": "sum",
            "Sell $": "sum",
        }
    )

    grouped_by_kit.sort_values(by="Sell $", ascending=False, inplace=True)

    new_df = pd.DataFrame(columns=df.columns)

    for index, row in grouped_by_kit.iterrows():
        new_df = pd.concat([new_df, df[df["Kit"] == index]])
        new_df = pd.concat(
            [
                new_df,
                pd.DataFrame(
                    [
                        [
                            "",
                            "Total",
                            row["Qty"],
                            "",
                            "",
                            "",
                            "",
                            "",
                            row["Sell $"],
                            "",
                            "",
                        ],
                        ["", "", "", "", "", "", "", "", "", "", ""],
                    ],
                    columns=df.columns,
                ),
            ]
        )

    total_qty = df["Qty"].sum()
    total_sell = df["Sell $"].sum()

    new_df = pd.concat(
        [
            new_df,
            pd.DataFrame(
                [
                    [
                        "",
                        "Grand Total",
                        total_qty,
                        "",
                        "",
                        "",
                        "",
                        "",
                        total_sell,
                        "",
                        "",
                    ]
                ],
                columns=df.columns,
            ),
        ]
    )

    formatted_total_sell = "${:,.2f}".format(total_sell)

    new_df.to_excel(
        f"Shortages2 Total-{formatted_total_sell} RunTime-{datetime.now(): %Y-%m-%d %H%M%S}.xlsx",
        index=False,
    )

    return df


def shortages3():
    global file_path

    df = pd.read_excel(file_path, sheet_name="Shortages3")

    adam_regex = re.compile(r"Adam", re.IGNORECASE)
    z_terr_regex = re.compile(r"23", re.IGNORECASE)
    chris_regex = re.compile(r"Chris", re.IGNORECASE)
    steve_regex = re.compile(r"Steve", re.IGNORECASE)
    brent_regex = re.compile(r"Brent", re.IGNORECASE)
    jeff_regex = re.compile(r"Jeff", re.IGNORECASE)
    dan_regex = re.compile(r"Dan", re.IGNORECASE)
    tom_regex = re.compile(r"Tom", re.IGNORECASE)
    rich_regex = re.compile(r"Rich", re.IGNORECASE)
    john_regex = re.compile(r"John", re.IGNORECASE)

    df.dropna(subset=["Sales Rep"], inplace=True)

    df["Sales Rep"] = df["Sales Rep"].astype(str)
    df["Sales Rep"] = df["Sales Rep"].str.replace("/", "-")

    df["Updated Unit Cost"] = df["Current Cost"] + df[0.03] + df[0.04]
    df["Total Updated Cost"] = df["Updated Unit Cost"] * df["Qty"]

    reps = df["Sales Rep"].unique()
    reps_dict = {}

    for rep in reps:
        if adam_regex.search(rep):
            reps_dict[rep] = "Adam Lichtenbaum"
        elif z_terr_regex.search(rep):
            reps_dict[rep] = "Adam Lichtenbaum"
        elif chris_regex.search(rep):
            reps_dict[rep] = "Rich Regruto"
        elif steve_regex.search(rep):
            reps_dict[rep] = "Steve Spicer"
        elif brent_regex.search(rep):
            reps_dict[rep] = "Brent Hill"
        elif jeff_regex.search(rep):
            reps_dict[rep] = "Jeff Wright"
        elif dan_regex.search(rep):
            reps_dict[rep] = "Dan Gildea"
        elif tom_regex.search(rep):
            reps_dict[rep] = "Tom Ranck"
        elif rich_regex.search(rep):
            reps_dict[rep] = "Rich Regruto"
        elif john_regex.search(rep):
            reps_dict[rep] = "John Casey"
        else:
            reps_dict[rep] = "House Account"

    df["Sales Rep Mapped"] = df["Sales Rep"].map(reps_dict)

    df["Sell $"] = df["Sell $"].astype(float).map("${:,.2f}".format)
    df["Qty"] = df["Qty"].astype(int).map("{:,}".format)

    for rep in reps:
        rep_df = df[df["Sales Rep Mapped"] == reps_dict[rep]][
            [
                "Kit",
                "Description",
                "Qty",
                "So Date",
                "Del Date",
                "SO Nbr",
                "Cust",
                "Cust Name",
                "Sell $",
                "Sales Rep",
            ]
        ]
        rep_df.sort_values(by=["Kit", "So Date"], inplace=True)
        rep_df.to_excel(f"{reps_dict[rep]}.xlsx", index=False)


def main():
    shortages1()
    shortages2()
    shortages3()


if __name__ == "__main__":
    main()
