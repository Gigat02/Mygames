# -*- coding: utf-8 -*-
"""Genera dati/giochi.xlsx a partire dalla collezione iniziale.

Script eseguito una volta sola per creare il file Excel di partenza.
Da qui in avanti il file viene aggiornato dalla web app (area Gestione).

    python tools/seed_excel.py
"""
import hashlib
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from descrizioni import DESCRIZIONI

# nome, min giocatori, max giocatori, durata min, durata max, tag
# (i riassunti stanno in tools/descrizioni.py)
GIOCHI = [
    ("Pictureka", 2, 6, 45, 45, "Party game, Family game, Osservazione, Bambini"),
    ("Forza 4", 2, 2, 20, 20, "Astratto, Per due, Classico, Family game"),
    ("Non ti arrabbiare", 2, 4, 60, 60, "Classico, Family game, Dadi, Tira e muovi"),
    ("Dixit", 3, 12, 30, 60, "Party game, Family game, Creatività, Gioco di carte"),
    ("Scythe", 1, 6, 120, 120, "Eurogame, Controllo territorio, Gestione risorse, Piazzamento lavoratori, Solitario, Gioco pesante"),
    ("Last Friday", 2, 6, 120, 120, "American, Horror, Movimento nascosto, Uno contro tutti, Deduzione"),
    ("Challengers! Beach Cup", 3, 8, 40, 40, "Deckbuilder, Gioco di carte, Family game, Torneo"),
    ("Smash Up", 2, 5, 40, 40, "Gioco di carte, Controllo territorio, American, Fantasy"),
    ("Las Vegas Party", 3, 12, 30, 60, "Dadi, Party game, Family game, Push your luck, Scommesse"),
    ("Twilight Imperium (quarta edizione)", 3, 8, 480, 480, "American, Fantascienza, Controllo territorio, Negoziazione, Gioco pesante, Epico"),
    ("Root", 2, 4, 120, 120, "Controllo territorio, Asimmetrico, Wargame, Gioco pesante, Fantasy"),
    ("Crescent Moon", 4, 5, 120, 120, "Controllo territorio, Asimmetrico, Negoziazione, Eurogame, Gioco pesante"),
    ("Zombicide: Dark Side", 2, 6, 30, 180, "Cooperativo, Miniature, American, Horror, Dadi, Solitario"),
    ("Zombicide: Season 1", 2, 12, 30, 180, "Cooperativo, Miniature, American, Horror, Dadi, Solitario"),
    ("In alto i calici", 2, 12, 30, 30, "Party game, Bluff, Memoria, Gioco di carte"),
    ("When I Dream", 4, 10, 30, 30, "Party game, Comunicazione limitata, A squadre, Deduzione"),
    ("Pozioni esplosive", 2, 4, 30, 30, "Eurogame, Family game, Puzzle, Set collection"),
    ("Camel Up", 2, 8, 45, 45, "Family game, Party game, Scommesse, Dadi, Push your luck"),
    ("Sagrada", 2, 4, 45, 45, "Eurogame, Family game, Dadi, Puzzle, Draft"),
    ("Whitechapel", 2, 6, 90, 90, "Movimento nascosto, Deduzione, Uno contro tutti, Storico"),
    ("Splendor", 2, 4, 30, 30, "Eurogame, Family game, Set collection, Motore di gioco"),
    ("Cryptid", 2, 5, 30, 30, "Deduzione, Investigazione, Family game, Logica"),
    ("7 Wonders", 3, 7, 30, 60, "Eurogame, Draft, Gioco di carte, Storico, Civilizzazione"),
    ("Capitan Sonar", 2, 8, 60, 60, "Party game, A squadre, Tempo reale, Deduzione"),
    ("Rock'n'Roll Robot", 1, 99, 30, 30, "Astratto, Puzzle, Logica, Solitario, Family game"),
    ("Wings of War", 2, 10, 30, 90, "Wargame, Storico, Gioco di carte, Simulazione"),
    ("Zombicide: Marvel Zombies - X-Men Resistance", 2, 12, 30, 180, "Cooperativo, Miniature, American, Horror, Dadi, Solitario"),
    ("Slay the Spire: il gioco da tavolo", 1, 4, 90, 270, "Deckbuilder, Cooperativo, Campagna, Solitario, Fantasy"),
    ("Memoir '44", 2, 2, 45, 45, "Wargame, Storico, Per due, Dadi, Gioco di carte"),
    ("Terra Nova", 2, 4, 45, 45, "Eurogame, Controllo territorio, Family game"),
    ("Trivial Pursuit", 2, 12, 60, 120, "Quiz, Party game, Family game, Classico, Cultura generale"),
    ("Avalon", 5, 10, 20, 30, "Party game, Ruoli nascosti, Bluff, Deduzione"),
    ("Legami di sangue", 6, 12, 30, 30, "Party game, Ruoli nascosti, Bluff, Deduzione"),
    ("Landmarks", 2, 10, 25, 25, "Party game, Parole, Cooperativo, Comunicazione limitata"),
    ("Carcassonne", 2, 6, 60, 120, "Eurogame, Family game, Piazzamento tessere, Classico"),
    ("Othello", 2, 2, 30, 30, "Astratto, Per due, Classico"),
    ("Il gioco delle uova", 2, 4, 30, 30, "Party game, Destrezza, Family game, Bambini"),
    ("Fuggi Fuggi!", 2, 7, 45, 45, "Family game, Corsa, Interazione cattiva, Horror"),
    ("Cash'n Guns", 4, 9, 30, 30, "Party game, Bluff, Negoziazione, Umoristico"),
    ("Ticket to Ride", 2, 6, 45, 45, "Eurogame, Family game, Set collection, Percorsi, Classico"),
    ("Similo", 2, 2, 15, 15, "Cooperativo, Deduzione, Gioco di carte, Filler"),
    ("Similo 2", 2, 2, 15, 15, "Cooperativo, Deduzione, Gioco di carte, Filler"),
    ("Il grande Dalmuti", 4, 12, 40, 60, "Gioco di carte, Party game, Family game, Grandi gruppi"),
    ("Tsuro", 2, 8, 25, 25, "Astratto, Family game, Piazzamento tessere, Filler"),
    ("Burrito", 2, 8, 20, 30, "Party game, Destrezza, Tempo reale, Family game, Umoristico"),
    ("Oh Issa!", 2, 2, 20, 20, "Gioco di carte, Per due, Filler, Family game"),
    ("Kahuna", 2, 2, 30, 30, "Eurogame, Per due, Controllo territorio, Gioco di carte"),
    ("King Up!", 2, 6, 30, 30, "Family game, Scommesse, Bluff, Eurogame"),
    ("Kamisado", 2, 2, 15, 15, "Astratto, Per due, Filler"),
    ("Kingdomino", 2, 4, 20, 30, "Family game, Piazzamento tessere, Draft, Eurogame"),
    ("7 Wonders Duel", 2, 2, 30, 30, "Eurogame, Per due, Draft, Gioco di carte, Storico"),
    ("Vudu", 3, 6, 30, 30, "Party game, Destrezza, Dadi, Umoristico"),
    ("Topiary", 2, 4, 25, 25, "Eurogame, Astratto, Family game, Piazzamento tessere"),
    ("Faraway", 2, 8, 20, 20, "Gioco di carte, Eurogame, Draft, Filler"),
    ("Just One", 4, 20, 20, 20, "Party game, Cooperativo, Parole, Family game, Grandi gruppi"),
    ("Cartagena", 2, 6, 30, 40, "Family game, Corsa, Gestione mano, Pirati"),
    ("3 Segreti", 2, 8, 20, 20, "Cooperativo, Deduzione, Investigazione, Party game"),
    ("Fort", 2, 4, 20, 20, "Deckbuilder, Gioco di carte, Eurogame, Family game"),
    ("Nome in codice", 4, 8, 20, 20, "Party game, Parole, A squadre, Deduzione"),
    ("Not Alone", 2, 7, 45, 45, "Uno contro tutti, Gioco di carte, Fantascienza, Bluff"),
    ("Guillotine", 2, 5, 30, 30, "Gioco di carte, Party game, Umoristico, Filler"),
    ("Bang!", 3, 8, 30, 40, "Ruoli nascosti, Gioco di carte, Party game, Western, Bluff"),
    ("Imagine", 3, 8, 30, 30, "Party game, Creatività, Deduzione"),
    ("Decripto", 3, 8, 30, 30, "Party game, A squadre, Deduzione, Parole"),
    ("Unstable Unicorns", 2, 8, 30, 60, "Gioco di carte, Party game, Umoristico, Interazione cattiva"),
    ("Tags", 2, 4, 30, 30, "Party game, Parole, Family game, A squadre"),
    ("Colt Express", 2, 6, 40, 40, "Family game, Programmazione, Western, Party game"),
    ("Mysterium", 2, 7, 40, 40, "Cooperativo, Deduzione, Comunicazione limitata, Horror"),
    ("Sator Arepo Tenet Opera Rotas", 2, 4, 60, 60, "Astratto, Eurogame, Gestione mano, Medievale"),
    ("MicroMacro", 1, 4, 45, 45, "Cooperativo, Investigazione, Osservazione, Family game, Solitario"),
    ("La casa delle mannaie", 2, 7, 30, 30, "Party game, Horror, Gioco di carte, Interazione cattiva, Umoristico"),
    ("Dungeon WC", 1, 5, 5, 5, "Cooperativo, Tempo reale, Party game, Filler, Umoristico"),
    ("Rock Paper Wizard", 3, 6, 30, 30, "Party game, Fantasy, Filler, Umoristico"),
    ("Rush & Bash", 2, 6, 30, 30, "Family game, Corsa, Fantasy, Interazione cattiva"),
    ("Otto minuti per un impero", 2, 5, 20, 20, "Eurogame, Controllo territorio, Filler"),
    ("Sky Team", 2, 2, 20, 20, "Cooperativo, Per due, Dadi, Comunicazione limitata, Puzzle"),
    ("Tristerra: La Grande Baldoria", 3, 8, 0, 0, "Party game, Narrativo, Fantasy, Gioco di ruolo"),
    ("Room 25", 1, 6, 30, 30, "Cooperativo, Ruoli nascosti, Fantascienza, Programmazione, Solitario"),
    ("Nome in codice: Visual", 4, 8, 20, 20, "Party game, A squadre, Deduzione, Osservazione"),
    ("The Resistance", 5, 10, 30, 30, "Party game, Ruoli nascosti, Bluff, Deduzione"),
    ("Lupusburg", 4, 8, 40, 40, "Party game, Ruoli nascosti, Bluff, Deduzione"),
    ("Happy Little Dinosaurs", 2, 4, 30, 60, "Gioco di carte, Party game, Umoristico, Family game"),
    ("Marvel Villainous", 2, 4, 30, 30, "Asimmetrico, Gioco di carte, American, Family game"),
    ("Pastiche", 2, 4, 60, 60, "Eurogame, Set collection, Piazzamento tessere, Arte"),
    ("Uno", 2, 6, 20, 30, "Gioco di carte, Family game, Classico, Filler"),
    ("Sei", 2, 10, 25, 25, "Gioco di carte, Family game, Push your luck, Filler, Classico"),
    ("Stay Away!", 4, 12, 30, 60, "Ruoli nascosti, Horror, Gioco di carte, Party game, Bluff"),
    ("Tesseract", 2, 4, 30, 30, "Cooperativo, Dadi, Puzzle, Solitario"),
    ("Mascarade", 2, 13, 30, 30, "Party game, Bluff, Ruoli nascosti, Memoria"),
    ("Intrighi a corte", 2, 5, 20, 20, "Gioco di carte, Bluff, Filler, Family game"),
    ("Timeline", 2, 6, 15, 15, "Family game, Quiz, Filler, Cultura generale"),
    ("Gloom", 2, 4, 60, 60, "Gioco di carte, Narrativo, Umoristico, Party game"),
    ("RRR", 2, 2, 10, 10, "Per due, Filler, Astratto"),
    ("Mangia Preda Chiama", 2, 5, 20, 20, "Party game, Bluff, Deduzione, Gioco di carte, Filler"),
    ("Korsar", 2, 8, 20, 20, "Gioco di carte, Family game, Filler, Pirati"),
    ("Warehouse 51", 3, 5, 45, 45, "Aste, Eurogame, Bluff"),
    ("Munchkin Cthulhu", 3, 6, 60, 60, "Gioco di carte, Umoristico, Interazione cattiva, Party game, Horror"),
    ("Zombie Dice", 2, 4, 10, 10, "Dadi, Push your luck, Filler, Party game"),
    ("One Zero One", 2, 2, 10, 10, "Astratto, Per due, Filler"),
    ("Venice Connection", 2, 2, 5, 5, "Astratto, Per due, Filler, Classico"),
    ("7 - The Sins", 2, 5, 15, 15, "Gioco di carte, Memoria, Filler, Push your luck"),
    ("Kluster", 1, 4, 10, 10, "Destrezza, Filler, Party game, Family game"),
    ("Niet!", 2, 5, 30, 30, "Gioco di carte, Prese, Eurogame, A squadre"),
    ("Orifiamma", 3, 5, 20, 20, "Gioco di carte, Bluff, Programmazione, Filler, Medievale"),
    ("Skull", 2, 12, 15, 15, "Party game, Bluff, Filler, Gioco di carte"),
    ("Cortex", 2, 6, 15, 15, "Party game, Family game, Rapidità, Bambini"),
    ("Match Up! Cose", 1, 10, 45, 45, "Party game, Cooperativo, Quiz, Family game, Solitario"),
    ("Match Up! Cibo", 1, 10, 45, 45, "Party game, Cooperativo, Quiz, Family game, Solitario"),
    ("7 Rosso", 2, 4, 10, 10, "Gioco di carte, Astratto, Filler"),
    ("The Game", 1, 5, 20, 20, "Cooperativo, Gioco di carte, Solitario, Filler"),
    ("I Cattivissimi 7", 2, 6, 25, 25, "Party game, Rapidità, Gioco di carte, Family game"),
    ("Coloretto", 2, 5, 30, 30, "Gioco di carte, Set collection, Family game, Filler"),
    ("Deckscape: Il destino di Londra", 1, 6, 90, 90, "Escape room, Cooperativo, Gioco di carte, Enigmi, Solitario"),
    ("Deckscape: Fuga da Alcatraz", 1, 6, 90, 90, "Escape room, Cooperativo, Gioco di carte, Enigmi, Solitario"),
    ("Deckscape: Ciurma contro ciurma", 1, 6, 90, 90, "Escape room, Cooperativo, Gioco di carte, Enigmi, Solitario"),
    ("Deckscape: Il mistero di Eldorado", 1, 6, 90, 90, "Escape room, Cooperativo, Gioco di carte, Enigmi, Solitario"),
    ("Deckscape: L'ora del test", 1, 6, 90, 90, "Escape room, Cooperativo, Gioco di carte, Enigmi, Solitario"),
    ("Burst", 3, 6, 15, 15, "Gioco di carte, Push your luck, Filler, Party game"),
    ("Blind Jack", 2, 12, 15, 15, "Party game, Quiz, Push your luck, A squadre, Family game"),
    ("Claim", 2, 2, 15, 15, "Gioco di carte, Prese, Per due, Filler, Fantasy"),
    ("Muffin Time", 2, 8, 30, 30, "Gioco di carte, Umoristico, Party game, Interazione cattiva"),
    ("Zombie Fluxx", 2, 6, 20, 40, "Gioco di carte, Party game, Umoristico, Regole variabili"),
    ("Paroliere", 2, 4, 20, 20, "Parole, Classico, Family game, Rapidità"),
    ("Minotaurus", 2, 4, 30, 60, "Family game, Bambini, Dadi, Labirinto"),
    ("Giorno di paga", 2, 5, 40, 40, "Family game, Classico, Tira e muovi, Economico"),
    ("Domino", 2, 2, 20, 20, "Astratto, Classico, Family game"),
    ("Radio Londra", 3, 6, 20, 20, "Party game, Deduzione, Bluff, A squadre, Comunicazione limitata"),
    ("Villa Paletti", 2, 4, 30, 30, "Destrezza, Family game, Costruzione"),
    ("Boon Lake", 1, 4, 90, 90, "Eurogame, Gestione risorse, Solitario, Gioco pesante"),
    ("Flick 'em Up!", 2, 10, 45, 45, "Destrezza, Family game, Western, A squadre"),
    ("Turing Machine", 1, 4, 20, 20, "Deduzione, Logica, Solitario, Eurogame"),
    ("Super Farmer", 2, 6, 30, 30, "Family game, Dadi, Bambini, Classico, Scambi"),
    ("Shit Happens", 2, 6, 30, 30, "Party game, Umoristico, Per adulti, Gioco di carte"),
    ("What Do You Meme?", 3, 10, 30, 30, "Party game, Umoristico, Per adulti"),
    ("Bad People", 3, 10, 30, 30, "Party game, Umoristico, Per adulti"),
    ("Hive", 2, 2, 20, 20, "Astratto, Per due, Classico"),
    ("Dobble", 2, 8, 5, 5, "Party game, Rapidità, Osservazione, Family game, Bambini, Filler"),
    ("Bananagrams", 1, 8, 20, 20, "Parole, Rapidità, Family game, Solitario"),
    ("Fantozzi - Batti lei", 2, 5, 20, 20, "Gioco di carte, Party game, Memoria, Umoristico, Family game"),
    ("Master Bluff", 3, 5, 20, 20, "Gioco di carte, Bluff, Filler, Family game"),
    ("Lupi Mannari", 8, 18, 30, 30, "Party game, Ruoli nascosti, Bluff, Deduzione, Grandi gruppi"),
    ("Il trauma del tram", 3, 13, 15, 15, "Party game, Umoristico, Per adulti, A squadre, Grandi gruppi"),
    ("Bag of Chips", 2, 5, 15, 15, "Push your luck, Filler, Family game"),
    ("Uno Golf", 1, 8, 20, 20, "Gioco di carte, Family game, Filler, Solitario"),
    ("Dadi poker", 1, 4, 30, 30, "Dadi, Classico, Push your luck, Family game, Solitario"),
    ("Perudo", 1, 6, 20, 20, "Dadi, Bluff, Party game, Classico"),
    ("Scacchi", 2, 2, 10, 10, "Astratto, Per due, Classico"),
]

# gioco base, espansione
ESPANSIONI = [
    ("Twilight Imperium (quarta edizione)", "L'ultimo tuono"),
    ("Twilight Imperium (quarta edizione)", "Profezia dei Re"),
    ("Twilight Imperium (quarta edizione)", "Codex 1, 2, 3, 4"),
    ("7 Wonders", "Cities"),
    ("7 Wonders", "Leaders"),
    ("Bang!", "Dodge City"),
    ("Bang!", "Terence Hill & Bud Spencer"),
    ("Memoir '44", "Eastern Front"),
    ("Wings of War", "6 aerei extra"),
    ("Zombicide: Marvel Zombies - X-Men Resistance", "Guardians of the Galaxy"),
    ("Zombicide: Marvel Zombies - X-Men Resistance", "Hydra Resurrection"),
    ("Zombicide: Marvel Zombies - X-Men Resistance", "The Sinister Six"),
    ("Zombicide: Season 1", "Toxic City Mall"),
    ("Zombicide: Season 1", "Zombie Dogs"),
    ("Unstable Unicorns", "Nightmare"),
    ("Timeline", "Celebrity"),
    ("Timeline", "Family"),
    ("Timeline", "Spazio (Samantha Cristoforetti)"),
    ("Timeline", "History"),
]

PASSWORD_INIZIALE = "mygames"

INTESTAZIONE = PatternFill("solid", fgColor="1F2430")
TESTO_INTESTAZIONE = Font(color="FFFFFF", bold=True)


def stile_foglio(ws, larghezze, nome_tabella):
    for indice, larghezza in enumerate(larghezze, start=1):
        ws.column_dimensions[get_column_letter(indice)].width = larghezza
    for cella in ws[1]:
        cella.fill = INTESTAZIONE
        cella.font = TESTO_INTESTAZIONE
        cella.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    riferimento = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    tabella = Table(displayName=nome_tabella, ref=riferimento)
    tabella.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight9", showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(tabella)


def main():
    wb = Workbook()

    giochi = wb.active
    giochi.title = "Giochi"
    giochi.append(["Nome", "Giocatori min", "Giocatori max", "Durata min", "Durata max",
                   "Tag", "Descrizione", "Note"])
    for nome, gmin, gmax, dmin, dmax, tag in sorted(GIOCHI, key=lambda g: g[0].lower()):
        giochi.append([nome, gmin, gmax, dmin or None, dmax or None, tag,
                       DESCRIZIONI.get(nome), None])
    stile_foglio(giochi, [42, 14, 14, 12, 12, 62, 90, 30], "TabellaGiochi")

    espansioni = wb.create_sheet("Espansioni")
    espansioni.append(["Gioco base", "Espansione", "Note"])
    for base, nome in ESPANSIONI:
        espansioni.append([base, nome, None])
    stile_foglio(espansioni, [42, 42, 30], "TabellaEspansioni")

    config = wb.create_sheet("Config")
    config.append(["Chiave", "Valore"])
    config.append(["password_hash", hashlib.sha256(PASSWORD_INIZIALE.encode()).hexdigest()])
    config.append(["versione", "2"])
    stile_foglio(config, [24, 72], "TabellaConfig")

    destinazione = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "dati", "giochi.xlsx")
    wb.save(destinazione)
    print(f"{len(GIOCHI)} giochi e {len(ESPANSIONI)} espansioni salvati in {destinazione}")


if __name__ == "__main__":
    main()
