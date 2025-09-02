import pandas as pd

def load_and_process_data(excel_file):
    df = pd.read_excel(excel_file, header=1, usecols="B:D")
    if df['Percent Active'].dtype == 'object':  
        df['Percent Active'] = df['Percent Active'].str.rstrip('%').astype(float)
    elif df['Percent Active'].max() <= 1.0:  
        df['Percent Active'] *= 100
    
    return df

def get_row_color(perc):
    if perc >= 70:
        return '#d4edda'  
    elif 60 <= perc < 70:
        return '#fff3cd'  
    else:
        return '#f8d7da'  

def generate_html(df, month, logo_mime, logo_base64):
    rows = ''
    for _, row in df.iterrows():
        color = get_row_color(row['Percent Active'])
        rows += f'''
        <tr style="background-color: {color};">
            <td style="padding: 5px; border: 1px solid #ddd;">{row['Rank']}</td>
            <td style="padding: 5px; border: 1px solid #ddd;">{row['Store']}</td>
            <td style="padding: 5px; border: 1px solid #ddd;">{row['Percent Active']:.2f}%</td>
        </tr>
        '''
    html = f'''
    <div style="font-family: Arial, sans-serif; max-width: 500px; margin: 0 auto; padding: 10px; border: 1px solid #ddd; background-color: #fff;">
        {'<img src="data:' + logo_mime + ';base64,' + logo_base64 + '" alt="Company Logo" style="height: 30px; vertical-align: middle; margin-left: 10px;">' if logo_base64 else ''}
        <h2 style="text-align: center; color: #7F27E4; margin-bottom: 5px; font-size: 18px;">{month} Leaderboard</h2>
        <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
            <thead>
                <tr style="background-color: #7F27E4; color: #fff; text-align: left;">
                    <th style="padding: 5px; border: 1px solid #ddd;">Rank</th>
                    <th style="padding: 5px; border: 1px solid #ddd;">Store</th>
                    <th style="padding: 5px; border: 1px solid #ddd;">% of Active Teammates</th>
                </tr>
            </thead>
            <tbody>
                {rows}
            </tbody>
        </table>
        
        <div style="margin-top: 5px; font-size: 10px; color: #333;">
            <p>Green: >= 70% active</p>
            <p>Yellow: 60%-70% active</p>
            <p>Red: <= 60% active</p>
        </div>
    </div>
    '''
    return html

if __name__ == "__main__":
    df = load_and_process_data('leaderboard-data.xlsx')
    html = generate_html(df, month='July')
    print(html)
