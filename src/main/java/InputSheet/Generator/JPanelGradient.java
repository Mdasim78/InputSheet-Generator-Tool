package InputSheet.Generator;

import java.awt.Color;
import java.awt.GradientPaint;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.RenderingHints;

import javax.swing.JPanel;

@SuppressWarnings("serial")
public class JPanelGradient extends JPanel {

private Color color1;
private Color color2;

protected JPanelGradient(Color color1, Color color2)
{
this.color1=color1;
this.color2=color2;
}
@Override
public void paintComponent(Graphics g){
super.paintComponent(g);
Graphics2D g2d = (Graphics2D)g.create();
int w = getWidth();
int h = getHeight();
g2d.setRenderingHint(RenderingHints.KEY_COLOR_RENDERING, RenderingHints.VALUE_COLOR_RENDER_QUALITY);
GradientPaint gp = new GradientPaint(
0, 0, color1,
0, h, color2);

g2d.setPaint(gp);
g2d.fillRect(0, 0, w, h);
}
}
