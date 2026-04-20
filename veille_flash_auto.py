import logging
import os
import platform
import shutil
import subprocess
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
from tavily import TavilyClient

TAVILY_API_KEY = "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"
DEFAULT_OUTPUT_NAME = "veille_flash"
THEMES_VEILLE = {
    "Concurrentielle": "concurrents et positionnement dans le secteur de l'archivage de donnees dans la sante",
    "Technique": "innovations techniques et avancees dans l'archivage de donnees",
    "Commerciale": "offres, marche, partenariats et strategie commerciale en archivage de donnees dans la sante",
    "Reglementaire": "reglementation, conformite et cadre legal de l'archivage de donnees dans la sante",
}


def determiner_dossier_sortie():
    dossier_env = os.getenv("VEILLE_FLASH_OUTPUT_DIR")
    if dossier_env:
        return Path(dossier_env).expanduser()

    home = Path.home()
    candidats = [
        home / "OneDrive" / "Projets Personnels",
        home / "Documents" / "Projets Personnels",
        home / "Projets Personnels",
        home / "veille_flash",
    ]
    for candidat in candidats:
        if candidat.parent.exists():
            return candidat
    return candidats[-1]


OUTPUT_DIR = determiner_dossier_sortie()
LOG_FILE = OUTPUT_DIR / "veille_flash_auto.log"
STATUS_FILE = OUTPUT_DIR / "dernier_statut_veille.txt"


def configurer_logs():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
        ],
        force=True,
    )


def charger_client_tavily():
    if not TAVILY_API_KEY:
        raise RuntimeError(
            "La cle Tavily est absente. Definis la variable d'environnement TAVILY_API_KEY."
        )
    return TavilyClient(api_key=TAVILY_API_KEY)


def normaliser_nom_feuille(valeur):
    caracteres_interdits = "[]:*?/\\"
    nom = "".join("_" if c in caracteres_interdits else c for c in valeur).strip()
    return (nom or DEFAULT_OUTPUT_NAME)[:31]


def generer_nom_sortie():
    horodatage = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return OUTPUT_DIR / f"veille_flash_hebdomadaire_{horodatage}.xlsx"


def ecrire_statut(statut, message, output_file=None, erreur_detail=None):
    horodatage = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lignes = [
        f"Statut : {statut}",
        f"Date : {horodatage}",
        f"Journal : {LOG_FILE}",
    ]
    if output_file:
        lignes.append(f"Rapport : {output_file}")
    lignes.append("")
    lignes.append(message)

    if erreur_detail:
        lignes.append("")
        lignes.append("Traceback :")
        lignes.append(erreur_detail)

    STATUS_FILE.write_text("\n".join(lignes), encoding="utf-8")


def notifications_desktop_activees():
    return os.getenv("NOTIFY_DESKTOP_ENABLED", "1").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }


def _echapper_powershell(texte):
    return texte.replace("`", "``").replace('"', '`"')


def _echapper_xml(texte):
    return (
        texte.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _envoyer_notification_windows(statut, sujet, message):
    sujet_ps = _echapper_powershell(_echapper_xml(sujet))
    message_ps = _echapper_powershell(_echapper_xml(message))
    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        raise RuntimeError("PowerShell est introuvable sur cette machine.")

    script = f"""
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] > $null

$xml = @"
<toast activationType="protocol" launch="file:///{STATUS_FILE.as_posix()}">
  <visual>
    <binding template="ToastGeneric">
      <text>{sujet_ps}</text>
      <text>{message_ps}</text>
    </binding>
  </visual>
</toast>
"@

$doc = New-Object Windows.Data.Xml.Dom.XmlDocument
$doc.LoadXml($xml)
$toast = [Windows.UI.Notifications.ToastNotification]::new($doc)
$notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Windows PowerShell")
$notifier.Show($toast)
"""
    logging.info("Envoi de la notification Windows %s", statut)
    subprocess.run(
        [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        check=True,
        timeout=20,
    )


def _echapper_applescript(texte):
    return texte.replace("\\", "\\\\").replace('"', '\\"')


def _envoyer_notification_macos(statut, sujet, message):
    osascript = shutil.which("osascript")
    if not osascript:
        raise RuntimeError("osascript est introuvable sur cette machine.")

    script = (
        f'display notification "{_echapper_applescript(message)}" '
        f'with title "{_echapper_applescript(sujet)}" '
        f'subtitle "{_echapper_applescript(f"Statut : {statut}")}"'
    )
    logging.info("Envoi de la notification macOS %s", statut)
    subprocess.run([osascript, "-e", script], check=True, timeout=20)


def _envoyer_notification_linux(statut, sujet, message):
    notify_send = shutil.which("notify-send")
    if not notify_send:
        raise RuntimeError("notify-send est introuvable sur cette machine.")

    urgence = "critical" if statut.lower() == "echec" else "normal"
    logging.info("Envoi de la notification Linux %s", statut)
    subprocess.run(
        [
            notify_send,
            "--app-name=Veille Flash",
            "--urgency",
            urgence,
            sujet,
            message,
        ],
        check=True,
        timeout=20,
    )


def envoyer_notification(statut, sujet, message):
    if not notifications_desktop_activees():
        logging.info("Notifications desktop desactivees.")
        return

    systeme = platform.system()
    try:
        if systeme == "Windows":
            _envoyer_notification_windows(statut, sujet, message)
        elif systeme == "Darwin":
            _envoyer_notification_macos(statut, sujet, message)
        elif systeme == "Linux":
            _envoyer_notification_linux(statut, sujet, message)
        else:
            logging.warning("Notifications non supportees sur %s", systeme)
    except Exception as exc:
        logging.warning("Notification desktop indisponible sur %s: %s", systeme, exc)


def recuperer_resultats_veille(tavily, type_veille, sujet):
    logging.info("=" * 50)
    logging.info("VEILLE FLASH : %s", type_veille.upper())
    logging.info("=" * 50)

    if not sujet:
        raise ValueError(f"Le sujet de veille est obligatoire pour {type_veille}.")

    query_complete = f"{sujet} actualites recentes et tendances. REPONDS EN FRANCAIS."

    logging.info("Analyse intelligente du web par Tavily...")
    search_result = tavily.search(
        query=query_complete,
        search_depth="advanced",
        topic="news",
        time_range="week",
        include_answer=True,
        max_results=10,
    )

    synthese_ia_globale = search_result.get(
        "answer",
        "Aucun resume n'a pu etre genere automatiquement.",
    )
    articles = search_result.get("results", [])

    logging.info("SYNTHESE IA GENEREE")
    logging.info("%s", synthese_ia_globale)

    resultats_data = []
    logging.info(
        "Generation des resumes individuels pour chaque site (cela peut prendre un moment)..."
    )

    for i, res in enumerate(articles, 1):
        url = res.get("url", "")
        titre = res.get("title", "Titre indisponible")
        logging.info("[%s/%s] Resume de : %s", i, len(articles), titre[:50])

        try:
            res_ia = tavily.search(
                query=f"Fais un resume synthetique en 2 phrases, EN FRANCAIS, de cet article : {url}",
                search_depth="advanced",
                topic="news",
                include_answer=True,
                max_results=1,
            )
            resume_final = res_ia.get("answer", "Resume non disponible.")
        except Exception as exc:
            logging.warning("Echec du resume pour %s: %s", url, exc)
            resume_final = "Erreur lors de la generation du resume."

        resultats_data.append(
            {
                "Type de veille": type_veille,
                "Sujet": sujet,
                "Titre": titre,
                "Resume Strategique": resume_final,
                "Score": res.get("score", 0),
                "Source": url,
            }
        )

    logging.info("Veille %s terminee", type_veille)
    return pd.DataFrame(resultats_data), synthese_ia_globale


def exporter_rapport(resultats_par_theme, output_file):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format(
            {"bold": True, "bg_color": "#004C97", "font_color": "white", "border": 1}
        )
        body_fmt = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
        score_fmt = workbook.add_format(
            {"num_format": "0.00", "align": "center", "valign": "top", "border": 1}
        )

        for type_veille, contenu in resultats_par_theme.items():
            feuille = normaliser_nom_feuille(type_veille)
            df = contenu["df"]
            synthese = contenu["synthese"]

            df.to_excel(writer, sheet_name=feuille, index=False, startrow=3)

            worksheet = writer.sheets[feuille]
            worksheet.write(0, 0, f"Type de veille : {type_veille}")
            worksheet.write(1, 0, f"Synthese globale : {synthese}")
            worksheet.set_column("A:A", 18, body_fmt)
            worksheet.set_column("B:B", 25, body_fmt)
            worksheet.set_column("C:C", 40, body_fmt)
            worksheet.set_column("D:D", 70, body_fmt)
            worksheet.set_column("E:E", 10, score_fmt)
            worksheet.set_column("F:F", 30, body_fmt)
            worksheet.set_row(3, 20, header_fmt)

    logging.info("Rapport pret : %s", output_file)
    return output_file


if __name__ == "__main__":
    configurer_logs()
    output_file = None

    try:
        tavily = charger_client_tavily()
        resultats_par_theme = {}

        for type_veille, sujet in THEMES_VEILLE.items():
            df, synthese = recuperer_resultats_veille(tavily, type_veille, sujet)
            resultats_par_theme[type_veille] = {"df": df, "synthese": synthese}

        output_file = exporter_rapport(resultats_par_theme, generer_nom_sortie())
        ecrire_statut(
            statut="SUCCES",
            message="Rapport de veille pret.",
            output_file=output_file,
        )
        envoyer_notification(
            statut="succes",
            sujet="Rapport de veille pret",
            message=(
                "Rapport de veille pret.\n\n"
                f"Fichier : {output_file}\n"
                f"Statut : {STATUS_FILE}"
            ),
        )
    except Exception as exc:
        erreur_detail = traceback.format_exc()
        logging.exception("Execution en echec: %s", exc)
        ecrire_statut(
            statut="ECHEC",
            message="Erreur du rapport de veille, consulter fichier statut de veille.",
            output_file=output_file,
            erreur_detail=erreur_detail,
        )

        try:
            envoyer_notification(
                statut="echec",
                sujet="Erreur du rapport de veille",
                message=(
                    "Erreur du rapport de veille, consulter fichier statut de veille.\n\n"
                    f"Statut : {STATUS_FILE}"
                ),
            )
        except Exception as notification_exc:
            logging.exception("Impossible d'envoyer la notification desktop: %s", notification_exc)

        raise SystemExit(1)
