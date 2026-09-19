# -*- coding: utf-8 -*-
"""Genera le icone PNG per la schermata Home a partire dal marchio dell'app.

Il marchio è quello disegnato in CSS nella testata (`.marchio__segno`):
un quadrato arrotondato color accento con un foro quadrato al centro.

    python tools/genera_icone.py
"""
import os

from PIL import Image, ImageDraw

SFONDO = (17, 17, 19, 255)        # --sfondo del tema scuro
ACCENTO = (255, 122, 77, 255)     # --accento del tema scuro

# nome file, lato, quota del marchio sul lato (le icone "maskable" lasciano
# margine perché Android ritaglia l'icona dentro forme diverse)
ICONE = [
    ('icona-180.png', 180, 0.60),
    ('icona-192.png', 192, 0.60),
    ('icona-512.png', 512, 0.60),
    ('icona-maskable-512.png', 512, 0.44),
]

SCALA = 4  # si disegna in grande e si rimpicciolisce, per bordi puliti


def disegna(lato, quota):
    grande = lato * SCALA
    immagine = Image.new('RGBA', (grande, grande), SFONDO)
    pennello = ImageDraw.Draw(immagine)

    marchio = grande * quota
    sinistra = (grande - marchio) / 2
    pennello.rounded_rectangle(
        [sinistra, sinistra, sinistra + marchio, sinistra + marchio],
        radius=marchio * 0.30,
        fill=ACCENTO,
    )

    foro = marchio * 0.40
    bordo = sinistra + (marchio - foro) / 2
    pennello.rounded_rectangle(
        [bordo, bordo, bordo + foro, bordo + foro],
        radius=foro * 0.25,
        fill=SFONDO,
    )

    return immagine.resize((lato, lato), Image.LANCZOS)


def main():
    cartella = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'assets')
    for nome, lato, quota in ICONE:
        destinazione = os.path.join(cartella, nome)
        disegna(lato, quota).save(destinazione, 'PNG', optimize=True)
        print(f'{nome}: {lato}×{lato}')


if __name__ == '__main__':
    main()
