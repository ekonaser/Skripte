import streamlit as st
import sqlite3
import pandas as pd
import plotly.express as px

# Funkcija za pridobitev podatkov iz SQLite baze
def fetch_data(query, params=()):
    conn = sqlite3.connect('databaseFDX.db')
    cursor = conn.cursor()
    cursor.execute(query, params)
    data = cursor.fetchall()
    conn.close()
    return data

# Pridobi podatke za zadnjih 10 dni in jih pretvori v DataFrame
query_dates = '''
SELECT Datum, COUNT(*) as count 
FROM obracuni_fdx 
WHERE Datum >= date('now', '-10 days') 
GROUP BY Datum
'''
data_dates = fetch_data(query_dates)
df_dates = pd.DataFrame(data_dates, columns=['Datum', 'count'])

# Pridobi podatke za PrejemnikPosiljatelj in jih pretvori v DataFrame
query_recipients = '''
SELECT PrejemnikPosiljatelj, COUNT(*) as count 
FROM obracuni_fdx 
GROUP BY PrejemnikPosiljatelj
'''
data_recipients = fetch_data(query_recipients)
df_recipients = pd.DataFrame(data_recipients, columns=['PrejemnikPosiljatelj', 'count'])

query_pogoji = '''
SELECT Pogoji, COUNT(*) as count 
FROM obracuni_fdx 
GROUP BY Pogoji
'''
data_pogoji = fetch_data(query_pogoji)
df_pogoji = pd.DataFrame(data_pogoji, columns=['Pogoji', 'count'])

# Pridobi vsoto vrednosti v stolpcu Garancija
query_sum_garancija = '''
SELECT SUM(Garancija) FROM obracuni_fdx 
'''
sum_garancija = fetch_data(query_sum_garancija)[0][0]

# Streamlit aplikacija
st.title('Obračuni FedEx')

# Prikaz interaktivnega stolpičnega grafa z uporabo Plotly za datume
st.subheader('Število poslanih obračunov za zadnjih 10 dni')
fig_dates = px.bar(df_dates, x='Datum', y='count', title='Število pošiljk za zadnjih 10 dni',
                   labels={'Datum': 'Datum', 'count': 'Število pošiljk'})

fig_dates.update_layout(
    xaxis=dict(title='Datum'),
    yaxis=dict(title='Število pošiljk'),
    clickmode='event+select'
)

st.plotly_chart(fig_dates)

# Prikaz tortnega grafa za Prejemnik/Posiljatelj
st.subheader('Porazdelitev pošiljk glede na Prejemnik/Pošiljatelj')
fig_recipients = px.pie(df_recipients, names='PrejemnikPosiljatelj', values='count', title='Porazdelitev pošiljk glede na Prejemnik/Posiljatelj')
st.plotly_chart(fig_recipients)

# Prikaz tortnega grafa za pogoje
st.subheader('Porazdelitev pošiljk glede na Pogoje')
fig_pogoji = px.pie(df_pogoji, names='Pogoji', values='count', title='Porazdelitev pošiljk glede na Pogoje')
st.plotly_chart(fig_pogoji)

# Prikaz vsote vrednosti v stolpcu Garancija v tabeli
st.subheader('Skupaj postavke')
table_data = {
    'Postavka': ['Garancija'],
    'Vrednost': [sum_garancija]
}
df_table = pd.DataFrame(table_data)
st.table(df_table.style.format({'Vrednost': '{:.2f}'}))
