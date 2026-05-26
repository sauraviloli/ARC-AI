import PptxGenJS from 'pptxgenjs';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { slides, title } = req.body;
  if (!slides || !Array.isArray(slides)) return res.status(400).json({ error: 'Invalid slides data' });

  try {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.title = title || 'Presentation';

    // Theme colors
    const DARK = '1a1915';
    const ACCENT = 'D97706';
    const LIGHT = 'f7f6f3';
    const MUTED = '6b6860';

    slides.forEach((slide, idx) => {
      const s = pptx.addSlide();

      // Background
      s.background = { color: idx === 0 ? DARK : 'FFFFFF' };

      if (idx === 0) {
        // TITLE SLIDE
        s.addShape(pptx.ShapeType.rect, { x: 0, y: 3.2, w: '100%', h: 0.06, fill: { color: ACCENT } });
        s.addText(slide.title || '', {
          x: 0.7, y: 1.2, w: 8.6, h: 1.4,
          fontSize: 40, bold: true, color: 'FFFFFF',
          fontFace: 'Calibri', align: 'left',
        });
        if (slide.subtitle) {
          s.addText(slide.subtitle, {
            x: 0.7, y: 2.7, w: 8.6, h: 0.7,
            fontSize: 20, color: 'cccccc',
            fontFace: 'Calibri', align: 'left',
          });
        }
        s.addText('Arc AI', {
          x: 0.7, y: 4.8, w: 3, h: 0.4,
          fontSize: 13, color: ACCENT, fontFace: 'Calibri',
        });
      } else {
        // CONTENT SLIDE
        // Top accent bar
        s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: ACCENT } });

        // Slide number
        s.addText(String(idx).padStart(2, '0'), {
          x: 8.8, y: 0.15, w: 0.8, h: 0.35,
          fontSize: 11, color: MUTED, align: 'right', fontFace: 'Calibri',
        });

        // Title
        s.addText(slide.title || '', {
          x: 0.5, y: 0.2, w: 8.2, h: 0.7,
          fontSize: 26, bold: true, color: DARK,
          fontFace: 'Calibri', align: 'left',
        });

        // Divider
        s.addShape(pptx.ShapeType.line, {
          x: 0.5, y: 1.0, w: 8.6, h: 0,
          line: { color: 'e5e4e0', width: 1 },
        });

        // Body bullets
        if (slide.bullets && slide.bullets.length) {
          const bulletItems = slide.bullets.map(b => ({
            text: b,
            options: { bullet: { type: 'bullet', indent: 15 }, fontSize: 17, color: '2c2c2a', paraSpaceAfter: 8 },
          }));
          s.addText(bulletItems, {
            x: 0.5, y: 1.15, w: 8.6, h: 3.4,
            fontFace: 'Calibri', valign: 'top',
          });
        } else if (slide.content) {
          s.addText(slide.content, {
            x: 0.5, y: 1.15, w: 8.6, h: 3.4,
            fontSize: 17, color: '2c2c2a',
            fontFace: 'Calibri', valign: 'top', wrap: true,
          });
        }

        // Footer
        s.addText((title || 'Presentation').toUpperCase(), {
          x: 0.5, y: 4.85, w: 6, h: 0.3,
          fontSize: 10, color: MUTED, fontFace: 'Calibri',
        });
      }
    });

    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    const filename = (title || 'presentation').toLowerCase().replace(/\s+/g, '-') + '.pptx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    return res.status(200).send(buffer);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'Failed to generate presentation: ' + err.message });
  }
}
