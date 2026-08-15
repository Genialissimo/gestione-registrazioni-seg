def _s21_centro_box(c: rl_canvas.Canvas, box: tuple, offset: float, testo: str = "X",
                     font_name: str = "Helvetica-Bold", font_size: float = 10.0, sposta: float = 0.0):
    x0, x1, top, bottom = box
    fattore_altezza_maiuscole = 0.717
    largo_testo = c.stringWidth(testo, font_name, font_size)
    x = (x0 + x1) / 2 - largo_testo / 2
    centro_verticale_top = (top + bottom) / 2 + offset
    baseline_top = centro_verticale_top + (font_size * fattore_altezza_maiuscole) / 2 + sposta
    y = S21_PAGE_H - baseline_top
    c.setFont(font_name, font_size)
    c.drawString(x, y, testo)
