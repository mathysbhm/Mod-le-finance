#Importer la librairie nécessaire
import pandas as pd
import matplotlib.pyplot as plt

#Charger le fichier Excel pour que Python puisse le lire
df = pd.read_excel(r"C:\Users\MATHIS\Desktop\Calculer des ratios python\entreprise.xlsx")

#Afficher les premières lignes
print("Données importées : ")
print(df.head)

#Calcul des ratios
df["Marge nette"] = df["Résultat net (€)"] / df["Chiffre d'affaires (€)"]
df["ROE"] = df["Résultat net (€)"] / df ["Capitaux propres (€)"]
df["Gearing"] = df["Dettes financières (€)"] / df ["Capitaux propres (€)"]

#Arrondir pour l'affichage
df["Marge nette"] = df ["Marge nette"].round(2)
df["ROE"] = df["ROE"].round(2)
df["Gearing"] = df["Gearing"].round(2)

# --- Étape projection financière ---

# 1. Calcul des taux de croissance annuels moyens
croissance_ca = (df["Chiffre d'affaires (€)"].iloc[-1] / df["Chiffre d'affaires (€)"].iloc[0]) ** (1 / (len(df)-1)) - 1
croissance_rn = (df["Résultat net (€)"].iloc[-1] / df["Résultat net (€)"].iloc[0]) ** (1 / (len(df)-1)) - 1
croissance_cp = (df["Capitaux propres (€)"].iloc[-1] / df["Capitaux propres (€)"].iloc[0]) ** (1 / (len(df)-1)) - 1
croissance_df = (df["Dettes financières (€)"].iloc[-1] / df["Dettes financières (€)"].iloc[0]) ** (1 / (len(df)-1)) - 1
croissance_ebitda = (df["EBITDA (€)"].iloc[-1] / df["EBITDA (€)"].iloc[0]) ** (1 / (len(df)-1)) - 1

# 2. Préparer les projections
annees_proj = [2024, 2025, 2026, 2027]
df_proj = pd.DataFrame({"Année": annees_proj})

# Initialiser avec les dernières valeurs connues
ca = df["Chiffre d'affaires (€)"].iloc[-1]
rn = df["Résultat net (€)"].iloc[-1]
cp = df["Capitaux propres (€)"].iloc[-1]
dfin = df["Dettes financières (€)"].iloc[-1]
ebitda = df["EBITDA (€)"].iloc[-1]

# Appliquer les croissances
for i in range(4):
    ca *= (1 + croissance_ca)
    rn *= (1 + croissance_rn)
    cp *= (1 + croissance_cp)
    dfin *= (1 + croissance_df)
    ebitda *= (1 + croissance_ebitda)

    df_proj.loc[i, "Chiffre d'affaires (€)"] = round(ca)
    df_proj.loc[i, "Résultat net (€)"] = round(rn)
    df_proj.loc[i, "Capitaux propres (€)"] = round(cp)
    df_proj.loc[i, "Dettes financières (€)"] = round(dfin)
    df_proj.loc[i, "EBITDA (€)"] = round(ebitda)

# 3. Calcul des ratios projetés
df_proj["Marge nette"] = df_proj["Résultat net (€)"] / df_proj["Chiffre d'affaires (€)"]
df_proj["ROE"] = df_proj["Résultat net (€)"] / df_proj["Capitaux propres (€)"]
df_proj["Gearing"] = df_proj["Dettes financières (€)"] / df_proj["Capitaux propres (€)"]

# 4. Fusionner avec les données initiales
df_full = pd.concat([df, df_proj], ignore_index=True)

# 5. Afficher les résultats projetés
print("\nProjection 2024–2027 :")
print(df_proj[["Année", "Chiffre d'affaires (€)", "Résultat net (€)", "Marge nette", "ROE", "Gearing"]])

# Créer un masque pour séparer réel et projection
df_full["Type"] = df_full["Année"].apply(lambda x: "Historique" if x <= 2023 else "Projection")

# Tracer
plt.figure(figsize=(10, 6))

# Historique
historique = df_full[df_full["Type"] == "Historique"]
plt.plot(historique["Année"], historique["Marge nette"], marker='o', label="Marge nette (réel)")
plt.plot(historique["Année"], historique["ROE"], marker='o', label="ROE (réel)")
plt.plot(historique["Année"], historique["Gearing"], marker='o', label="Gearing (réel)")

# Projection
projection = df_full[df_full["Type"] == "Projection"]
plt.plot(projection["Année"], projection["Marge nette"], marker='o', linestyle='--', label="Marge nette (projection)")
plt.plot(projection["Année"], projection["ROE"], marker='o', linestyle='--', label="ROE (projection)")
plt.plot(projection["Année"], projection["Gearing"], marker='o', linestyle='--', label="Gearing (projection)")

# Habillage
plt.title("Évolution des ratios financiers : Historique vs Projection")
plt.xlabel("Année")
plt.ylabel("Ratio")
plt.grid(True)
plt.legend()
plt.tight_layout()
plt.show()

#Exporter le fichier dans un nouveau excel
df_full.to_excel("rapport financier complet.xlsx", index=False)

#Vérifier la bonne éxécution du script
print("\n Le script s'est éxécuté de manière normale.")
