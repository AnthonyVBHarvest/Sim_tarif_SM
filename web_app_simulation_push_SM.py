import pandas as pd
import seaborn as sns
import streamlit as st
import fonctions_fact as ff

df = pd.read_csv("Data_O2S_since_2021_SM.csv")

df_old_fact = pd.read_csv("Data_facture_old_SM.csv")



############################################## Streamlite app



#################################### Setting



st.set_page_config(page_title="Simulation tarifaire O2S",
                   page_icon=":chart_with_upwards_trend:",
                   layout= "wide")



onglets = st.sidebar.radio('Onglets', 
                 options=["Simulation"])


# Select Values


def select_values():

    global df
    global df_old_fact
    global palier
    global prix_palier
    global prix_palier_s
    global encours_min
    global imposition_encours
    global df_new_price
    global df_old_fact
    global liste_col_new_fact_clients_palier
    global date_choix
    global type_client_fact
    global type_c_f
    global type_c_f_no_s
    global nb_ged_offert
    global prix_ged
    global palette_2x
    global signature_prix
    global MP_PM_prix
    global MP_std_prix


    ## Paramètre facturation Pourcentage de contrat O2S
    st.markdown("<h2 style='text-align: center; color: black;'>Paramètres Facturation \n</h2>", unsafe_allow_html=True)

    # Choix Date
    date_choix = st.selectbox("Quelle plage temporelle choisir", ["2022", "2021"])

    df = df.loc[(df["Date"].str.contains(date_choix))]
    df_old_fact = df_old_fact.loc[(df_old_fact["Date"].str.contains(date_choix))]

    # Choix contrat date de signature
    left, right = st.columns(2)
    date_signature_choix_min = left.text_input("Date de signature de contrat minimum (YYYY-MM)", value="2007-01")
    date_signature_choix_max = right.text_input("Date de signature de contrat maximun (YYYY-MM)", value=df["Date de signature"].max())

    df = df.loc[(df["Date de signature"]>=date_signature_choix_min) & (df["Date de signature"]<=date_signature_choix_max)]
    df_old_fact = df_old_fact.loc[(df_old_fact["Date de signature"]>=date_signature_choix_min) & (df_old_fact["Date de signature"]<=date_signature_choix_max)]

    ## Parmètre facturation client
    st.markdown("<h4 style='text-align: center; color: black;'>Paramètres Facturation Client  \n</h4>", unsafe_allow_html=True)

    # Pourcentage de contrat O2S
    type_client_fact = st.selectbox("Quelle base pour la facturation client ?", ["Nombre de comptes agrégés", "Total des clients et prospects agrégés"])
    if type_client_fact == "Total des clients et prospects agrégés":
        type_c_f = "clients"
        type_c_f_no_s = "client"
    else: 
        type_c_f = "comptes"
        type_c_f_no_s = "compte"

    fig_hist_distri_clients, fig_boxplot_clients, fig_distrib_ponderer = ff.visualisation_distrib(df, type_client_fact, date_choix)

    st.markdown(f"<h4 style='text-align: center; color: black;'>Distribution des entreprises par le {type_client_fact} \n</h4>", unsafe_allow_html=True)
    left, right = st.columns(2)
    left.plotly_chart(fig_hist_distri_clients)
    right.plotly_chart(fig_boxplot_clients)
    left.plotly_chart(fig_distrib_ponderer)

    percent_sucess = st.number_input("Nombre de contrat O2S Sucess (en % du total proportionnellement réparti par paliers)", value=20, max_value=100, min_value=0)

    

    # Palier en nombre de clients

    nb_paliers = st.number_input("Combien de paliers voulez-vous?", value=4, step=1)

    palier_nb_client_par_defaut = [50, 300, 500, 1000, 2000, 3000, 4000, 5000] + [5000 + 1000 * (x+1) for x in range(nb_paliers-8) if nb_paliers > 8]
    palier_prix_par_defaut = [0.70, 0.60, 0.50, 0.40, 0.35, 0.30, 0.25, 0.20] + [0.2 - 0.01 * (x+1) for x in range(nb_paliers-8) if nb_paliers > 8]
    palier_prix_s_par_defaut = [0.85, 0.75, 0.65, 0.55, 0.45, 0.40, 0.35, 0.30] + [0.3 - 0.01 * (x+1) for x in range(nb_paliers-8) if nb_paliers > 8]

    palier_nb_client_par_defaut = palier_nb_client_par_defaut[:nb_paliers]
    palier_prix_par_defaut = palier_prix_par_defaut[:nb_paliers]
    palier_prix_s_par_defaut = palier_prix_s_par_defaut[:nb_paliers]

    min_val_pr = 0.12

    left, center, right = st.columns(3)

    palier_0 = left.number_input("Palier 0", min_value=0, step=0)
    prix_palier_0 = center.number_input('Prix groupé du palier de base (en €)', value=140.0, min_value=min_val_pr)
    prix_s_palier_0 = right.number_input('Prix Sucess groupé du palier de base (en €)', value=160.0, min_value=min_val_pr)

    for i in range(len(palier_nb_client_par_defaut)):
        globals()[f"palier_{i+1}"] = left.number_input(f'Palier {i+1} - Nombre limite pour la base choisie', value=palier_nb_client_par_defaut[i])
        globals()[f"prix_palier_{i+1}"] = center.number_input(f'Palier {i+1} - Prix par base choisie HT (en €)', value=palier_prix_par_defaut[i], min_value=min_val_pr)
        globals()[f"prix_s_palier_{i+1}"] = right.number_input(f'Palier {i+1} - Prix Sucess par base choisie HT (en €)', value=palier_prix_s_par_defaut[i], min_value=min_val_pr)

    # Autres paramètres
    st.markdown("<h4 style='text-align: center; color: black;'>Paramètres Facturation Licence  \n</h4>", unsafe_allow_html=True)
    left, right = st.columns(2)

    prix_licence = left.number_input('Prix par licence (en €)', value=70)

    nb_acces_offert = right.number_input("Nombre d'accès offert avec le palier 0", value=2, min_value=0)

    st.markdown("<h4 style='text-align: center; color: black;'>Paramètres Facturation GED  \n</h4>", unsafe_allow_html=True)
    left, right = st.columns(2)

    prix_ged = left.number_input('Prix par Go (en €)', value=5.0)

    nb_ged_offert = right.number_input("Nombre de Go GED avec le palier 0", value=2.0, min_value=0.0)

    st.markdown("<h4 style='text-align: center; color: black;'>Paramètres Facturation Encours  \n</h4>", unsafe_allow_html=True)
    left, right = st.columns(2)

    encours_min = left.number_input('Encours moyens minimal pour facturation (en €)', value=100000)
    imposition_encours= right.number_input("Imposition de l'encours (en % / 1000)", value=0.0)

    st.markdown("<h4 style='text-align: center; color: black;'>Paramètre Facturation Signature  \n</h4>", unsafe_allow_html=True)
    signature_prix = st.number_input('Prix par signature', value=2.0)

    st.markdown("<h4 style='text-align: center; color: black;'>Paramètres Facturation MoneyPitch  \n</h4>", unsafe_allow_html=True)
    left, right = st.columns(2)
    MP_PM_prix = left.number_input('Prix par MoneyPitch Premium', value=5.0)
    MP_std_prix = right.number_input('Prix par MoneyPitch Standard', value=0.5)

    st.markdown("<h2 style='text-align: center; color: black;'>----------------------------------------------------------------------------------------------  \n\n\n</h2>", unsafe_allow_html=True)

    

    # Définition des listes-paliers
    palier = [palier_0] + [globals()[f"palier_{i+1}"] for i in range(len(palier_nb_client_par_defaut))] + [1e100]

    prix_palier = [prix_palier_0] + [globals()[f"prix_palier_{i+1}"] for i in range(len(palier_prix_par_defaut))]

    prix_palier_s = [prix_s_palier_0] + [globals()[f"prix_s_palier_{i+1}"] for i in range(len(palier_prix_s_par_defaut))]

    liste_col_new_fact_clients_palier = []
    for i in range(len(palier)-2):
        liste_col_new_fact_clients_palier.append(f"Facture O2S New Client - Palier {palier[i+1]}")
    liste_col_new_fact_clients_palier.append(f"Facture O2S New Client - Palier > {palier[-2]}")

    for i in range(len(palier)-2):
        liste_col_new_fact_clients_palier.append(f"Facture O2S-Sucess New Client - Palier {palier[i+1]}")
    liste_col_new_fact_clients_palier.append(f"Facture O2S-Sucess New Client - Palier > {palier[-2]}")


    df_new_price = ff.nouveau_prix_O2S_et_sucess(
                                    df, 
                                    palier, 
                                    prix_palier,
                                    prix_palier_s, 
                                    percent_sucess,
                                    type_client_fact,
                                    prix_licence,
                                    nb_acces_offert,
                                    nb_ged_offert,
                                    prix_ged,
                                    encours_moyen_seuil=encours_min, 
                                    imposition_encours=imposition_encours,
                                    signature_prix=signature_prix,
                                    MP_PM_prix=MP_PM_prix,
                                    MP_std_prix=MP_std_prix)

    palette_2x = [sns.color_palette("PuBu").as_hex()[1]]*(len(palier)-1) +[sns.color_palette("YlOrBr").as_hex()[1]]*(len(palier)-1)


# Liste

col_visu_split_sucess_normal=[
    "Facture O2S-Normal New Client",
    "Facture O2S-Sucess New Client"
]

liste_col_split_fact = [
    'Facture O2S Client', 
    'Facture O2S Licence',
    'Facture O2S Encours',
    "Facture O2S GED", 
    "Facture O2S Signature", 
    "Facture O2S MoneyPitch"
]

liste_new_col_split_fact = [
    'Facture O2S New Client',
    'Facture O2S New Licence',
    'Facture O2S New Encours',
    "Facture O2S New GED",
    "Facture O2S New Signature", 
    "Facture O2S New MoneyPitch"
]

liste_col_export_base = [
    "N° de contrat",
    "Société : Nom de la société",
    "Date",
    "Total des clients et prospects agrégés",
    "O2S - Nombre d'accès total",
    "Encours agrégé en €",
    "GED - Nombre d'unités utilisées",
    "Signature@ - Nombre de signatures à facturer",
    "Nombre d'accès MoneyPitch ouverts",
    "Nombre d'accès MoneyPitch Premium ouverts",
    "Nombre de connexions moyennes par semaine par utilisateur",
]

#################################### PAGE



def Simulation(df_new_price, df_old_fact):



    fig_compar_bar = ff.comparaison_bar(df_old_fact, df_new_price, ["Facture O2S"], ["Facture O2S"], f"Comparaison de la facturation O2S-SM en {date_choix}")
                                
    fig_compar_split_bar = ff.comparaison_bar(df_old_fact, df_new_price, liste_col_split_fact, liste_new_col_split_fact, f"Comparaison de la facturation splitté de O2S-SM en {date_choix}")
    
    fig_fact_split = ff.fig_facture_split(df_new_price, liste_col_new_fact_clients_palier, f"Simulation de la nouvelle facturation client de O2S-SM par palier en {date_choix}", palette_2x)

    fig_encours = ff.fig_x_y_evo_encours(df_new_price, imposition_encours)
    
    fig_imposition_effet = ff.fig_x_y_evo_encours_imposition(df_new_price, encours_min, imposition_encours)

    fig_effet_centime = ff.effet_changement_prix_palier(df_new_price, palier, prix_palier, prix_palier_s, type_client_fact, palette_2x)

    fig_part_sucess = ff.pie_part_sucess(df_new_price, date_choix)

    fig_part_split = ff.pie_rep_spli(df_new_price, liste_new_col_split_fact, date_choix)


    # Mise en forme
    st.markdown("<h2 style='text-align: center; color: black;'>Évaluation globale \n</h2>", unsafe_allow_html=True)
    left, right = st.columns(2)
    left.plotly_chart(fig_compar_bar)
    right.plotly_chart(fig_compar_split_bar)
    left.plotly_chart(fig_part_split)

    st.markdown("<h2 style='text-align: center; color: black;'>Évaluation Facturation client  \n</h2>", unsafe_allow_html=True)
    left, right = st.columns(2)
    left.plotly_chart(fig_fact_split)
    right.plotly_chart(fig_part_sucess)



    st.markdown("<h2 style='text-align: center; color: black;'>Évaluation de changements  \n</h2>", unsafe_allow_html=True)
    left, right = st.columns(2)
    left.plotly_chart(fig_imposition_effet)
    right.plotly_chart(fig_encours)
    left.plotly_chart(fig_effet_centime)


    ## Download Excel
    df_old_fact = df_old_fact.rename(columns={"Facture O2S": "Facturation théorique actuelle"})[liste_col_export_base+["Facturation théorique actuelle"]]
    df_new_price = df_new_price.rename(columns={"Facture O2S": "Facturation théorique proposée"})[liste_col_export_base+["Facturation théorique proposée"]]
    df_export = df_new_price.merge(df_old_fact, how="left", on=liste_col_export_base)
    df_xlsx = ff.to_excel(df_export)

    st.download_button(
    "Appuyez pour télécharger vos données",
    data = df_xlsx,
    file_name= 'df_test.xlsx')




#################################### Mise en page



if onglets == "Simulation":
    st.title("Simulation SM")
    select_values()
    Simulation(df_new_price, df_old_fact)





















