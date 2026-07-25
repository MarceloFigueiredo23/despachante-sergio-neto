# -*- coding: utf-8 -*-
"""Gera 4 criativos 1080x1080 + DOCX para Google Docs."""
from __future__ import annotations

import subprocess
import time
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
OUT = Path(__file__).resolve().parent
LOGO = ROOT / "logo.png"
EDGE = Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")

POSTS = [
    {
        "slug": "01-site",
        "dia": "Dia 1 — publicar hoje",
        "eyebrow": "COSMÓPOLIS · SP",
        "title": "AGORA TEMOS<br>SITE",
        "body": "Conheça nossos serviços e fale direto pelo WhatsApp. Transferência, licenciamento, IPVA e mais.",
        "cta": "Link na bio e no Google",
        "accent": "site",
        "copy": (
            "Agora temos site!\n"
            "Conheça nossos serviços e fale direto pelo WhatsApp:\n"
            "🌐 https://marcelofigueiredo23.github.io/despachante-sergio-neto/\n"
            "Instagram: @despachante_sergioneto\n"
            "Despachante em Cosmópolis — transferência, licenciamento, IPVA e mais."
        ),
    },
    {
        "slug": "02-licenciamento",
        "dia": "Dia 4 — daqui 3–4 dias",
        "eyebrow": "NÃO DEIXE PARA DEPOIS",
        "title": "LICENCIAMENTO<br>EM DIA",
        "body": "No Despachante Sérgio Neto (Biju) a gente cuida da documentação com atendimento próximo.",
        "cta": "Chama no WhatsApp 19 99185-1849",
        "accent": "lic",
        "copy": (
            "Já pensou no licenciamento do seu veículo?\n"
            "Não deixe para a última hora. No Despachante Sérgio Neto (Biju) a gente cuida da documentação com atendimento próximo em Cosmópolis.\n"
            "📍 Av. Éster, 119 — Bela Vista III\n"
            "📲 Chama no WhatsApp e tire sua dúvida."
        ),
    },
    {
        "slug": "03-transferencia",
        "dia": "Semana 2",
        "eyebrow": "COMPROU OU VENDEU?",
        "title": "TRANSFERÊNCIA<br>SEM DOR DE CABEÇA",
        "body": "Orientamos os documentos certos e acompanhamos o processo até a regularização.",
        "cta": "Fale conosco no WhatsApp",
        "accent": "transf",
        "copy": (
            "Comprou ou vendeu um carro?\n"
            "A transferência precisa estar correta para não dar dor de cabeça depois.\n"
            "Aqui orientamos os documentos certos e acompanhamos o processo.\n"
            "Despachante Sérgio Neto — Cosmópolis-SP\n"
            "Fale conosco e resolva sem enrolação."
        ),
    },
    {
        "slug": "04-documentos",
        "dia": "Semana 2 — após o post 3",
        "eyebrow": "CHECKLIST RÁPIDO",
        "title": "DOCUMENTOS DA<br>TRANSFERÊNCIA",
        "body": "RG/CNH · CRLV · Comprovante de endereço · Recibo de compra e venda. Confirme conosco antes.",
        "cta": "Tire dúvida no WhatsApp",
        "accent": "docs",
        "copy": (
            "Documentos mais pedidos na transferência:\n"
            "✅ Documento do comprador e do vendedor\n"
            "✅ CRLV / documentação do veículo\n"
            "✅ Comprovante de endereço\n"
            "✅ Recibo de compra e venda preenchido\n"
            "Lista pode variar — confirme conosco antes de ir ao cartório/DETRAN.\n"
            "Despachante Sérgio Neto | Cosmópolis"
        ),
    },
]


def html_for(post: dict) -> str:
    logo_uri = LOGO.as_uri()
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8" />
<style>
  @font-face {{ font-family: ImpactLocal; src: local('Impact'); }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html, body {{ width: 1080px; height: 1080px; overflow: hidden; }}
  body {{
    font-family: "Segoe UI", Arial, sans-serif;
    color: #eef3f9;
    background:
      radial-gradient(ellipse 70% 55% at 85% 15%, rgba(232,163,23,.28), transparent 55%),
      radial-gradient(ellipse 50% 40% at 10% 90%, rgba(90,140,210,.22), transparent 50%),
      linear-gradient(145deg, #2a4a78 0%, #1a3358 48%, #12263f 100%);
    position: relative;
  }}
  body::before {{
    content: "";
    position: absolute; inset: 0;
    background-image: repeating-linear-gradient(-18deg, transparent, transparent 42px, rgba(238,243,249,.035) 42px, rgba(238,243,249,.035) 43px);
  }}
  .frame {{
    position: relative; z-index: 1;
    width: 1080px; height: 1080px;
    padding: 64px 68px 56px;
    display: flex; flex-direction: column;
  }}
  .top {{
    display: flex; align-items: center; gap: 22px;
    margin-bottom: 48px;
  }}
  .logo {{
    width: 148px; height: 148px; border-radius: 50%;
    background: #fff;
    box-shadow: 0 0 0 4px rgba(232,163,23,.55), 0 18px 40px rgba(0,0,0,.35);
  }}
  .brand-line {{
    font-size: 22px; font-weight: 700; letter-spacing: .16em;
    text-transform: uppercase; color: rgba(238,243,249,.7);
  }}
  .eyebrow {{
    display: inline-flex; align-items: center; gap: 14px;
    font-size: 22px; font-weight: 700; letter-spacing: .18em;
    color: #e8a317; margin-bottom: 22px;
  }}
  .eyebrow::before {{ content: ""; width: 42px; height: 3px; background: #e8a317; }}
  h1 {{
    font-family: ImpactLocal, Impact, "Arial Black", sans-serif;
    font-size: 92px; line-height: .95; letter-spacing: .02em;
    font-weight: 400; margin-bottom: 28px; max-width: 12ch;
    text-transform: uppercase;
  }}
  .body {{
    font-size: 30px; line-height: 1.4; font-weight: 500;
    color: rgba(238,243,249,.88); max-width: 16ch; margin-bottom: auto;
  }}
  .footer {{
    display: flex; align-items: center; justify-content: space-between;
    border-top: 1px solid rgba(238,243,249,.18);
    padding-top: 28px; margin-top: 36px;
  }}
  .cta {{
    display: inline-flex; align-items: center; gap: 14px;
    background: #25d366; color: #062312;
    font-weight: 800; font-size: 24px;
    padding: 18px 28px; border-radius: 999px;
  }}
  .wa {{
    width: 34px; height: 34px; border-radius: 50%;
    background: #062312; color: #25d366;
    display: grid; place-items: center; font-size: 18px; font-weight: 800;
  }}
  .handle {{
    text-align: right; font-size: 20px; font-weight: 600;
    color: rgba(238,243,249,.55); letter-spacing: .04em;
  }}
  .num {{
    position: absolute; right: 48px; top: 220px;
    font-family: ImpactLocal, Impact, sans-serif;
    font-size: 220px; line-height: 1; color: rgba(232,163,23,.12);
    letter-spacing: -.04em; pointer-events: none;
  }}
</style>
</head>
<body>
  <div class="frame">
    <div class="num">{post['slug'][:2]}</div>
    <div class="top">
      <img class="logo" src="{logo_uri}" alt="" />
      <div class="brand-line">Despachante Sérgio Neto · Biju</div>
    </div>
    <div class="eyebrow">{post['eyebrow']}</div>
    <h1>{post['title']}</h1>
    <p class="body">{post['body']}</p>
    <div class="footer">
      <div class="cta"><span class="wa">W</span>{post['cta']}</div>
      <div class="handle">@despachante_sergioneto</div>
    </div>
  </div>
</body>
</html>
"""


def screenshot(html_path: Path, png_path: Path) -> None:
    png_path.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        str(EDGE),
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        f"--screenshot={png_path}",
        "--window-size=1080,1080",
        "--default-background-color=00000000",
        html_path.as_uri(),
    ]
    subprocess.run(cmd, check=True, capture_output=True)
    # Edge sometimes needs a beat on slow disks
    for _ in range(20):
        if png_path.exists() and png_path.stat().st_size > 1000:
            break
        time.sleep(0.2)
    if not png_path.exists():
        raise RuntimeError(f"Screenshot failed: {png_path}")


def build_docx(paths: list[tuple[dict, Path]]) -> Path:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")

    h = doc.add_heading("Criativos prontos — Despachante Sérgio Neto", level=0)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1A, 0x33, 0x58)

    p = doc.add_paragraph()
    r = p.add_run(
        "Use este arquivo no Google Docs (Arquivo → Abrir / Upload). "
        "Em cada post: baixe/salve a imagem → publique no Google Meu Negócio e Instagram com a copy."
    )
    r.italic = True
    r.font.size = Pt(10)

    doc.add_paragraph("Canais: Google Meu Negócio + Instagram @despachante_sergioneto")
    doc.add_paragraph("WhatsApp: (19) 99185-1849")
    doc.add_paragraph(
        "Site: https://marcelofigueiredo23.github.io/despachante-sergio-neto/"
    )
    doc.add_paragraph()

    for post, img in paths:
        hh = doc.add_heading(f"{post['slug'].upper()} — {post['dia']}", level=1)
        for run in hh.runs:
            run.font.color.rgb = RGBColor(0x1A, 0x33, 0x58)

        cb = doc.add_paragraph()
        cb.add_run("☐  ").bold = True
        cb.add_run("Publicado no Google — data: __________")
        cb2 = doc.add_paragraph()
        cb2.add_run("☐  ").bold = True
        cb2.add_run("Publicado no Instagram — data: __________")

        doc.add_paragraph("Criativo (imagem):")
        doc.add_picture(str(img), width=Inches(4.8))
        last = doc.paragraphs[-1]
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER

        lab = doc.add_paragraph()
        r = lab.add_run("Copy pronta (copiar e colar na legenda):")
        r.bold = True
        r.font.color.rgb = RGBColor(0x1A, 0x33, 0x58)

        copy_p = doc.add_paragraph(post["copy"])
        copy_p.paragraph_format.space_after = Pt(18)
        doc.add_paragraph("— — — — — — — — — — — — — — —")

    out = OUT / "posts-criativos-despachante.docx"
    doc.save(out)
    return out


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    results: list[tuple[dict, Path]] = []
    for post in POSTS:
        html_path = OUT / f"{post['slug']}.html"
        png_path = OUT / f"{post['slug']}.png"
        html_path.write_text(html_for(post), encoding="utf-8")
        print("html", html_path.name)
        screenshot(html_path, png_path)
        print("png", png_path.name, png_path.stat().st_size)
        results.append((post, png_path))

    docx = build_docx(results)
    print("docx", docx)

    # Galeria HTML simples
    cards = []
    for post, img in results:
        cards.append(
            f"""
            <section class="card">
              <h2>{post['slug']} · {post['dia']}</h2>
              <img src="{img.name}" alt="{post['slug']}" width="540" />
              <pre>{post['copy']}</pre>
              <p>☐ Google &nbsp;&nbsp; ☐ Instagram &nbsp;&nbsp; data: ______</p>
            </section>
            """
        )
    gallery = OUT / "galeria-posts.html"
    gallery.write_text(
        f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8" />
<title>Galeria de posts</title>
<style>
body{{font-family:Arial,sans-serif;max-width:720px;margin:2rem auto;padding:0 1rem;color:#1a3358}}
img{{max-width:100%;border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.12)}}
pre{{background:#eef3f9;padding:1rem;white-space:pre-wrap;border-left:4px solid #e8a317}}
.card{{margin:2rem 0;padding-bottom:1.5rem;border-bottom:1px solid #dce6f2}}
</style></head><body>
<h1>Criativos + copy — Despachante Sérgio Neto</h1>
<p>Abra o DOCX no Google Docs ou baixe as PNGs da pasta <code>criativos/</code>.</p>
{''.join(cards)}
</body></html>
""",
        encoding="utf-8",
    )
    print("gallery", gallery)


if __name__ == "__main__":
    main()
