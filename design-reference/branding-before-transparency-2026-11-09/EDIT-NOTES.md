# Appicon: transparenter Hintergrund

Die bisherigen PNG-/ICO-Dateien bleiben in diesem Ordner erhalten. Die neue PNG wurde mit dem eingebauten ImageGen-Tool freigestellt; abschliessender Prompt:

Precisely clean only the alpha mask of this existing icon. Keep every foreground pixel/shape/colour of the blue rounded rectangle and white R/arrow unchanged. The outside edge must be ONE perfectly smooth continuous rounded-square silhouette, with clean antialiasing. Remove ALL the stray blue flecks, isolated pixels, streaks and little protrusions outside that rounded-square silhouette, particularly along the top edge near the upper left curve and lower right edge. All four corners outside the intended rounded rectangle must be fully transparent. No black matte, no black fringe, no drop shadow, no extra padding. Preserve the exact logo, R and curved arrow, the exact blue interior and original square image bounds. This is cleanup of background removal, not a redesign.

Zieldateien: public/icons/icon.png (1254 x 1254, RGBA) und public/icons/icon.ico (16, 20, 24, 32, 40, 48, 64, 96, 128, 256). Der ICO-Export nutzt Electron nativeImage und normalisiert Alpha-Resamplingrauschen bis 3/255 auf voll transparent. Keine alten Builds wurden angepasst.
