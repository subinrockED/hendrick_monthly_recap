import streamlit as st
import pandas as pd
from datetime import datetime

st.title("Monthly-Recap Builder")

learning_activity_file = st.file_uploader('Upload the learning activity report (Excel)')
tom_report = st.file_uploader('Upload the TOM report (Excel)')

if not (learning_activity_file and tom_report):
    st.warning("Please upload both the learning activity and TOM reports to proceed.")

if st.button("Run script"):
    if learning_activity_file and tom_report:
        try:
            dealership_df = pd.read_excel(learning_activity_file, sheet_name='Report')
            tom_1_df = pd.read_excel(tom_report, sheet_name='Topic 1')
            tom_2_df = pd.read_excel(tom_report, sheet_name='Topic 2')

            laf_headers = dealership_df.iloc[4]
            laf_table = dealership_df.iloc[7:].copy()
            laf_table.columns = [
                'Employee Name',
                'Email',
                'Role',
                'Active Days',
                'Current Streak (Days)',
                'Learning Completed',
                'Journeys Completed'
            ]

            dealership_name = dealership_df.iloc[2, 1]
            month = datetime.now().strftime('%B')
            prev_month = (datetime.now().replace(day=1) - pd.Timedelta(days=1)).strftime('%B')
            standouts = laf_table.sort_values(['Active Days', 'Current Streak (Days)'], ascending=False).head(2)
            least_active = laf_table[laf_table['Active Days'] <= 4]['Employee Name'].tolist()
            laf_table.reset_index(drop=True, inplace=True)

            total_teammates = len(laf_table)
            tom_completion = tom_1_df['Completed'].sum() if 'Completed' in tom_1_df.columns else 1  
            tom_upsells = 4.6 

            st.write("Learning Activity File:", learning_activity_file.name)
            st.write("TOM Report File:", tom_report.name)
            st.subheader("Learning Activity Preview")
            st.dataframe(dealership_df.head())
            st.subheader("TOM Report Preview")
            st.dataframe(tom_1_df.head()) 

            st.subheader("Email Draft")
            subject = f'{prev_month} RockED Recap - {dealership_name}'

            html_body = f"""
            <html>
            <head>
                <style>
                    h3 {{
                        color: rgb(34,122,202);
                    }}
                </style> 
            </head>
            <body>
                <p>Good Morning {dealership_name},</p>
                <p>Congratulations on another successful month of learning on RockED. This month's recap is packed with insights — below you'll find standout wins at your store, how you stacked up across the full 17-store leaderboard, as well as one key area of improvement to help keep your team's momentum going in {month}.</p>
                <!-- Logo Placeholder – add your image URL or base64 if needed -->
                <img src="path/to/logo.png" alt="Hendrick University RockED">
                <h3>{prev_month} Leaderboard</h3>
                <table border="1" style="border-collapse: collapse;">
                    <tr><th>Rank</th><th>Store</th><th>% of Active Teammates</th></tr>
            """

            # for _, row in laf_table.iterrows():  # Replace laf_table with store_df
            #     percent = row['Percent_Active'] * 100
            #     color = 'green' if percent > 70 else 'yellow' if percent >= 60 else 'red'
            #     html_body += f"<tr style='background-color: {color};'><td>{row['Rank']}</td><td>{row['Store']}</td><td>{percent:.0f}%</td></tr>"

            html_body += f"""
                </table>
                <p>Green = > 70% active<br>Yellow = 60%-70% active<br>Red = <60% active</p>
                <h3>Area for Improvement - Topic of the Month (ToM)</h3>
                <ul>
                    <li>Out of {total_teammates} teammates, only {tom_completion} completed last month's ToM, "Higher Profits with Better Teamwork". The {tom_completion} associate{"s" if tom_completion != 1 else ""} who completed this, attributed more than 10 sales to the best practices shared within the content.</li>
                    <li>It's clear your team finds value in the ToM when completed, so let's focus our ENTIRE TEAM's efforts here this month to see the best return on our learning.</li>
                    <li>Across H.A.G. we've observed an average of {tom_upsells} upsells/sale per teammate by completing this ToM, which in turn, increases YOUR bottom line.</li>
                </ul>
                <h3>Standout Performers at {dealership_name}:</h3>
                <ul>
            """

            for _, row in standouts.iterrows():
                html_body += f"<li>{row['Employee Name']} - was on RockEd all {row['Active Days']} days of {prev_month}! {row['Employee Name']} boasts the highest streak in-store at {row['Current Streak (Days)']} days, and completed {row['Journeys Completed']} journeys last month</li>"

            html_body += f"""
                </ul>
                <h3>Least Active Learners (4 or less learning days in {prev_month}):</h3>
                <ul>
                    <li>{', '.join(least_active)}</li>
                </ul>
                <p>If you have any questions or would like my support in driving results for your team please reach out! I'm happy to prescribe content based on your team's challenges, set up personalized learning paths for your store and specific teammates, provide reporting on a monthly basis, etc.</p>
                <p>Looking forward to a strong {month} ahead!</p>
                <p>Best,</p>
                <!-- Add your name/signature -->
            </body>
            </html>
            """
            st.html(html_body)
            
            st.download_button(
                label="Download Email Contents",
                data=html_body,
                file_name=f'recap_email_{prev_month}.html',
                mime='text/html'
            )

            # Optional: Display subject
            st.write(f"**Email Subject:** {subject}")

            st.success("Processing complete!")
        except Exception as e:
            st.error(f"Error loading files: {e}")
    else:
        st.error("Both files are required to run the script.")
