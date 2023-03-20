import numpy as np
import pandas as pd
import plotly.graph_objs as go
import plotly_express as px
from scipy import stats
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

# Taille graphique

largeur = 750
hauteur = 600


#################################### Fonction prix 

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
    "Facture O2S Signatures", 
    "Facture O2S MoneyPitch"
]

liste_new_col_split_fact = [
    'Facture O2S New Client',
    'Facture O2S New Licence',
    'Facture O2S New Encours',
    "Facture O2S New GED",
    "Facture O2S New Signatures", 
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

#Nouveau prix - Version Optimisé

def nouveau_prix_O2S_et_sucess(df, palier, prix_palier, prix_palier_s, percent_sucess, type_client_fact,
                             prix_licence, nb_acces_offert, nb_ged_offert, prix_ged,
                             encours_moyen_seuil, imposition_encours, signature_prix,
                             MP_PM_prix, MP_std_prix):
      
    df1 = df.copy()  
    
    df1 = df1.sort_values(by=["Date", f"{type_client_fact}"])
    
    #Facturation par client
    
    df1["Facture O2S client"] = 0
    df1["Catégeorie palier client"] = 0
    
    for i in range(len(palier)-1):
        df1[f"Facture O2S client - palier {i}"] = 0
        df1[f"Facture O2S-Sucess client - palier {i}"] = 0
    
    for i in range(len(palier)-1):
        df1["Catégeorie palier client"].loc[(df1[f"{type_client_fact}"]>palier[i]) & (df1[f"{type_client_fact}"]<=palier[i+1])] = i
            
        if percent_sucess > 0:
            j = np.arange(df1.loc[df1["Catégeorie palier client"]==i].shape[0], step=(100/percent_sucess))
            j = [round(z,0) for z in j]
            df2 = df1.copy()
            df2 = df2.loc[df2["Catégeorie palier client"]==i].sort_values(by=["Date", f"{type_client_fact}"]).reset_index(drop=True)
            df2["Catégeorie palier client"].loc[(df2["Catégeorie palier client"]==i)&(df2.index.isin(j))] = i +10
            df1 = df1.drop(df1.loc[df1["Catégeorie palier client"]==i].index, axis=0)
            df1 = pd.concat([df1, df2], axis=0, ignore_index=True)
    
    for prix_pal, terme in zip([prix_palier, prix_palier_s], [0,10]):
        sucess_name=""
        if terme == 10:
            sucess_name = "-Sucess"
        
        #Catégorie 0
        cat = 0 + terme
        df1["Facture O2S client"].loc[df1["Catégeorie palier client"]==cat] = prix_pal[0]
        df1[f"Facture O2S{sucess_name} client - palier 0"].loc[df1["Catégeorie palier client"]==cat] = prix_pal[0]

        for i in range(len(palier)-2):
            cat = i+1 + terme 
            df1["Facture O2S client"].loc[df1["Catégeorie palier client"]==cat] = prix_pal[0] + sum([((palier[k+2] - palier[k+1])*prix_pal[k+1]) for k in range(i)]) + \
                ((df1[f"{type_client_fact}"].loc[df1["Catégeorie palier client"]==cat]-palier[i+1])*prix_pal[i+1])
            df1[f"Facture O2S{sucess_name} client - palier 0"].loc[df1["Catégeorie palier client"]==cat] = prix_pal[0]
            for z in range(i):
                df1[f"Facture O2S{sucess_name} client - palier {z+1}"].loc[df1["Catégeorie palier client"]==cat] = ((palier[z+2]-palier[z+1])*prix_pal[z+1]) 
            df1[f"Facture O2S{sucess_name} client - palier {i+1}"].loc[df1["Catégeorie palier client"]==cat] = ((df1[f"{type_client_fact}"].loc[df1["Catégeorie palier client"]==cat]-palier[1])*prix_pal[1])

    for i in range(len(palier)-1):
        df1[f"Facture O2S client - palier {i} (en % du CA client)"] = df1[f"Facture O2S client - palier {i}"] / df1["Facture O2S client"] * 100
        df1[f"Facture O2S-Sucess client - palier {i} (en % du CA client)"] = df1[f"Facture O2S-Sucess client - palier {i}"] / df1["Facture O2S client"] * 100
    

    #Facturation par licence
    df1["Facture O2S Utilisateur"] = 0
    df1["Facture O2S Utilisateur"].loc[df1["O2S - Nombre d'accès total"]>nb_acces_offert] = (df1["O2S - Nombre d'accès total"]-nb_acces_offert) * prix_licence   

    #Facturation par l'encours
    df1["Facture O2S Encours"] = 0
    df1["Facture O2S Encours"].loc[df1["Encours moyen par clients agrégés et prospects agrégés"]>encours_moyen_seuil] = (df1["Encours agrégé en €"] - (df1["Total des clients et prospects agrégés"]*encours_moyen_seuil)) * (imposition_encours/1000) / 100
    
    #Facturation GED
    df1["Facture O2S GED"] = 0
    df1["Facture O2S GED"].loc[df1["GED - Nombre d'unités utilisées"]>nb_ged_offert] = np.ceil((df1["GED - Nombre d'unités utilisées"]-nb_ged_offert)) * prix_ged 

    #Facturation Signatures
    df1["Facture O2S Signature"] = 0
    df1["Facture O2S Signature"] = df1["Signature@ - Nombre de signatures à facturer"]*signature_prix

    
    #Facturation MoneyPitch
    df1["Facture O2S MoneyPitch"] = 0
    df1["Facture O2S MoneyPitch"] =  df1["Nombre d'accès MoneyPitch Premium ouverts"]*MP_PM_prix + df1["Nombre d'accès MoneyPitch ouverts"]*MP_std_prix


    #Facturation totale
    df1["Facture O2S"] = df1["Facture O2S client"] + df1["Facture O2S Utilisateur"] + df1["Facture O2S Encours"] + df1["Facture O2S GED"] + \
                         df1["Facture O2S Signature"] + df1["Facture O2S MoneyPitch"]
    
    
    #Renommage
    df1 = df1.rename(columns={
    "Facture O2S client":"Facture O2S New Client",
    "Facture O2S Utilisateur":"Facture O2S New Licence",
    "Facture O2S Encours":"Facture O2S New Encours",
    "Facture O2S GED":"Facture O2S New GED",
    "Facture O2S Signature":"Facture O2S New Signature",
    "Facture O2S MoneyPitch":"Facture O2S New MoneyPitch"})

    for i in range(len(palier)-1):
        for terme in ["", "-Sucess"]:
            df1 = df1.rename(columns={
                f"Facture O2S{terme} client - palier {i}": f"Facture O2S{terme} New Client - Palier {palier[i+1]}",
                f"Facture O2S{terme} client - palier {i} (en % du CA client)": f"Facture O2S{terme} Client New - Palier {palier[i+1]} (en % du CA client)",
                })

            df1 = df1.rename(columns={
                f"Facture O2S{terme} New Client - Palier 1e+100": f"Facture O2S{terme} New Client - Palier > {palier[-2]}",
                f"Facture O2S{terme} New Client - Palier 1e+100 (en % du CA client)": f"Facture O2S{terme} New Client - Palier > {palier[-2]} (en % du CA client)",
                })

    df1[f"Facture O2S-Normal New Client"] = np.sum(df1[list(df1.columns[(df1.columns.str.contains("Facture O2S New Client - Palier")) & (~df1.columns.str.contains("-Sucess"))& (~df1.columns.str.contains("en %"))])], axis=1)
    df1[f"Facture O2S-Sucess New Client"] = np.sum(df1[list(df1.columns[(df1.columns.str.contains("-Sucess"))& (~df1.columns.str.contains("en %"))])], axis=1)
    
    
    return df1

########## Fonction 

# Fonction hors visu

def separateur_millier(df, col):
    
    df[f"{col} (arrondi)"] = df[col].apply(lambda x: round(x,0))
    df[f"{col} (arrondi)"] = df[f"{col} (arrondi)"].astype("int64")
    df[f"{col} (sep)"] = df[f"{col} (arrondi)"].apply(lambda x: format(x, ',d'))
    df[f"{col} (sep)"] = df[f"{col} (sep)"].str.replace(","," ")
    del df[f"{col} (arrondi)"]
    
    return df[f"{col} (sep)"]



def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data



# Fonction visu

def comparaison_bar(df_ancien, df_nouveau, liste_col_ancien, liste_col_nouveau, titre):
    
    df1_ancien = df_ancien[liste_col_ancien].sum().to_frame().T
    df1_nouveau = df_nouveau[liste_col_nouveau].sum().to_frame().T
    
    for i in range(len(liste_col_ancien)):
        df1_nouveau = df1_nouveau.rename(columns={f"{liste_col_nouveau[i]}":f"{liste_col_ancien[i]}"})
    
    
    for col in liste_col_ancien:
        df1_ancien[f"{col} (sep)"] = separateur_millier(df1_ancien, col)
        df1_nouveau[f"{col} (sep)"] = separateur_millier(df1_nouveau, col)
    
    
    variation_comp = [round((df1_nouveau[x] / df1_ancien[x] - 1)*100, 1) for x in liste_col_ancien]
    
    int_fig_1 = go.Bar(
        x = liste_col_ancien,
        y = df1_ancien[liste_col_ancien].sum(),
        name = "Ancien",
        textposition = "inside",
        text = [df1_ancien[f"{x} (sep)"][0] + " €" for x in liste_col_ancien]
        ) 
    
    int_fig_2 = go.Bar(
        x = liste_col_ancien,
        y = df1_nouveau[liste_col_ancien].sum(),
        name = "Nouveau",
        textposition = "inside",
        text = [str(df1_nouveau[f"{x} (sep)"][0]) + " €" + '<br> ' + str(v[0]) + "%" for x, v in zip(liste_col_ancien, variation_comp)]
        )
    
    data_stack_hist = [int_fig_1,
                       int_fig_2]
    
    layout_stack_hist = go.Layout(barmode ="group",
                                  title = titre,
                                  legend=dict(yanchor="top", y=1.2, xanchor="left", x=0.8))
    
    fig = go.Figure(
        data = data_stack_hist,
        layout = layout_stack_hist)
    
    return fig



def fig_facture_split(df, liste_col, titre, palette_2x) :
    
    df1 = df[liste_col].sum()
    
    df1_text = df1.to_frame().T
    
    for col in liste_col:
        df1_text[f"{col} (sep)"] = separateur_millier(df1_text, col)
    
    fig = go.Figure()
    
    fig = px.bar(x=liste_col,
                 y=df1,
                text=[df1_text[f"{x} (sep)"][0] for x in liste_col],
                color=palette_2x,
                color_discrete_map="identity")
    
    fig.update_layout(title=titre,
                        width=largeur, 
                        height=hauteur)
    
    return fig



def fig_x_y_evo_encours(df, imposition_encours, encours_moyen_seuil_plage = 400000):
    
    df1 = df.copy()
    
    df_test_1 = pd.DataFrame()
    df_test_1["Seuil encours"] = [i for i in range(0, encours_moyen_seuil_plage, 10000)]
    
    liste_effet_seuil = []
    for seuil in df_test_1["Seuil encours"]:
        df_99 = df1.copy()
        df_99["Facture O2S Encours"] = 0
        df_99["Facture O2S Encours"].loc[df_99["Encours moyen par clients agrégés et prospects agrégés"]>seuil] = (df_99["Encours agrégé en €"] - (df_99["Total des clients et prospects agrégés"]*seuil)) * (imposition_encours/1000) / 100
        liste_effet_seuil.append(df_99["Facture O2S Encours"].sum())
    

    df_test_1["Facture encours"] = liste_effet_seuil
    
    fig= px.line(df_test_1,
                x="Seuil encours",
                y="Facture encours",
                title = f"Facturation de l'encours en fonction du seuil de minimun de l'encours moyen pour une imposition de {imposition_encours} %",
                width=largeur,
                height=hauteur)
    
    return fig    


    
def fig_x_y_evo_encours_imposition(df, encours_moyen_min, imposition_encours):
    
    df1 = df.copy()
    
    df_test_1 = pd.DataFrame()
    df_test_1["Imposition encours"] = [i for i in np.arange(0, imposition_encours+2, 0.5)]
    
    liste_effet_seuil = []
    for seuil in df_test_1["Imposition encours"]:
        df_99 = df1.copy()
        df_99["Facture O2S Encours"] = 0
        df_99["Facture O2S Encours"].loc[df_99["Encours moyen par clients agrégés et prospects agrégés"]>encours_moyen_min] = (df_99["Encours agrégé en €"] - (df_99["Total des clients et prospects agrégés"]*encours_moyen_min)) * (seuil/1000) / 100
        liste_effet_seuil.append(df_99["Facture O2S Encours"].sum())
    

    df_test_1["Facture encours"] = liste_effet_seuil
    
    fig= px.line(df_test_1,
                x="Imposition encours",
                y="Facture encours",
                title = f"Facturation de l'encours en fonction de l'imposition de {imposition_encours}/1000 % et d'un encours moyen min de {encours_moyen_min} €",
                width=largeur,
                height=hauteur)
    fig.update_xaxes(title='Imposition encours / 1000')

    return fig    



def pie_part_sucess(df, date_choix):
    
    df1 = df.copy()
    df1 = df1[col_visu_split_sucess_normal].sum().reset_index()
    
    fig = px.pie(
       df1, 
       values=round(df1[0],0),
       names="index",
       title = f"Répartition de la facturation client entre O2S Sucess et normal en {date_choix}",
       width=largeur,
       height=hauteur)

    fig.update_traces(textinfo='value+percent',
                     marker=dict(colors=["silver", "yellow"], line=dict(color='#000000', width=2)))
    
    
    return fig 



def pie_rep_spli(df, liste_col, date_choix):
    
    df1 = df[liste_col].sum().reset_index()
    
    df1 = df1.rename(columns={"index":"Nom",
                        0:"Valeur"})
    
    fig = px.pie(df1, 
                values = round(df1["Valeur"], 0),
                names = "Nom",
                title = f"Répartition par split pour la nouvelle facturation O2S-SM en {date_choix}",
                color_discrete_sequence=px.colors.qualitative.Pastel,
                width=largeur,
                height=hauteur)
    fig.update_traces(textinfo='value+percent', marker=dict(line=dict(color='#000000', width=1)))

    return fig



def visualisation_distrib(df, type_client_fact, date_choix):

    df1 = df.copy()
    df1 = df1.loc[df1["Date"].str.contains(f"{date_choix}-12")]

    fig = px.histogram(df1[f"{type_client_fact}"].loc[df1[f"{type_client_fact}"]<10000], nbins=10,
            title=f"Distribution des établissements SM en fonction du {type_client_fact}",
            labels={"count":"Nombre d'entreprises",
                    "value":f"Taille des entreprises en {type_client_fact}"},
                    width=largeur,
                    height=hauteur)
    fig.update_yaxes(title="Nombre d'entreprises")
    fig.update(layout_showlegend=False)

    liste_nb_client = []
    for i in np.arange(10000, step=1000):
        nb_client = df1[f"{type_client_fact}"].loc[(df1[f"{type_client_fact}"] >= i) & \
                                                                                            (df1[f"{type_client_fact}"] < (i+1000))].sum()
        liste_nb_client.append(nb_client)
        
    fig_pondere = px.bar(x=[f"{x}-{x+999}" for x in np.arange(10000, step=1000)],
                         y=liste_nb_client,
                         labels={"x":f"Taille des entreprises en {type_client_fact}",
                                 "y":f"{type_client_fact}"},
                         width=largeur,
                         height=hauteur,
                         title = f"{type_client_fact} par tranche d'entreprise en fonction de leur taille de {type_client_fact}")  

    fig1 = px.box(df1[f"{type_client_fact}"],
                width=largeur,
                height=hauteur)
    fig1.update_yaxes(title=f"{type_client_fact}")


    return fig, fig1, fig_pondere



def effet_changement_prix_palier(df, palier, prix_palier, prix_palier_s, type_client_fact, palette_2x):
    
    df1=df.copy()
        
    liste_seuil_N = []
    liste_seuil_S = []

    liste_effet_N = []
    liste_effet_S = []

    liste_effet_par_euro_N =[]
    liste_effet_par_euro_S =[]


    for n in range(len(prix_palier)):
        df_test_1=pd.DataFrame()
        
        if n == 0:
            df_test_1[f"Seuil palier {n}"] = [i for i in np.arange(prix_palier[n]-1, prix_palier[n]+1, 1)]
        else:
            df_test_1[f"Seuil palier {n}"] = [i for i in np.arange(prix_palier[n]-0.01, prix_palier[n]+0.01, 0.01)]
            
        if n == 0:
            df_test_1[f"Seuil Sucess palier {n}"] = [i for i in np.arange(prix_palier_s[n]-1, prix_palier_s[n]+1, 1)]
        else:
            df_test_1[f"Seuil Sucess palier {n}"] = [i for i in np.arange(prix_palier_s[n]-0.01, prix_palier_s[n]+0.01, 0.01)]
        
        liste_seuil_N.append(df_test_1[f"Seuil palier {n}"])
        liste_seuil_S.append(df_test_1[f"Seuil Sucess palier {n}"])
        
        for ll, pal in enumerate([df_test_1[f"Seuil palier {n}"], df_test_1[f"Seuil Sucess palier {n}"]]): 
            liste_effet_palier_N=[]
            liste_effet_palier_S=[]

            for pa in pal:
                
                prix_d=prix_palier.copy()
                prix_d_s=prix_palier_s.copy()

                if ll == 0:
                    prix_d[n]=pa
                else:
                    prix_d_s[n]=pa

                df1["Facture O2S client"] = 0

                #Catégorie normal
                cat = 0 
                df1["Facture O2S client"].loc[df1["Catégeorie palier client"]==cat] = prix_d[0]

                for i in range(len(prix_palier)-1):
                    cat = i+1 
                    df1["Facture O2S client"].loc[df1["Catégeorie palier client"]==cat] = prix_d[0] + sum([((palier[k+2] - palier[k+1])*prix_d[k+1]) for k in range(i)]) + \
                        ((df1[f"{type_client_fact}"].loc[df1["Catégeorie palier client"]==cat]-palier[i+1])*prix_d[i+1])

                #Catégorie Sucess
                cat = 10 
                df1["Facture O2S client"].loc[df1["Catégeorie palier client"]==cat] = prix_d_s[0]

                for i in range(len(prix_palier)-1):
                    cat = i+11 
                    df1["Facture O2S client"].loc[df1["Catégeorie palier client"]==cat] = prix_d_s[0] + sum([((palier[k+2] - palier[k+1])*prix_d_s[k+1]) for k in range(i)]) + \
                        ((df1[f"{type_client_fact}"].loc[df1["Catégeorie palier client"]==cat]-palier[i+1])*prix_d_s[i+1])
                
                if ll ==0:
                    liste_effet_palier_N.append(df1["Facture O2S client"].sum())
                else:
                    liste_effet_palier_S.append(df1["Facture O2S client"].sum())
        
            if ll == 0:
                liste_effet_N.append(liste_effet_palier_N)
                LR = stats.linregress(liste_seuil_N[n], liste_effet_N[n])
                liste_effet_par_euro_N.append(round(LR[0]/100, 0))
            else:
                liste_effet_S.append(liste_effet_palier_S)
                LR = stats.linregress(liste_seuil_S[n], liste_effet_S[n])
                liste_effet_par_euro_S.append(round(LR[0]/100, 0))
    
    liste_effet_par_euro = liste_effet_par_euro_N + liste_effet_par_euro_S
    
    x_label = [(x) for x in [str(n0) + "- "+ str(n1) for n0, n1 in zip(palier[:-1],palier[1:])]] + \
                [(x) for x in [str(n0) + "- "+ str(n1) + "Sucess" for n0, n1 in zip(palier[:-1],palier[1:])]]
    
    
    fig=px.bar(x=x_label,
               y=liste_effet_par_euro,
               text= liste_effet_par_euro,
               title="Effet de l'augmentation de 1 centime pour chaque palier de prix",
               labels={"x":"Paliers",
                        "y":"Augmentation (en €)"},
               color=palette_2x,
               color_discrete_map="identity",
                width=largeur,
                height=hauteur)
    
    return fig
