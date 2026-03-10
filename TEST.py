import subprocess
import os

tikz_code = r"""
\documentclass{standalone}
\usepackage{tikz}
\begin{document}
\begin{array}{cc}
\text{A} & \text{B} \\
\begin{tikzpicture}
\draw[fill = blue!20, rotate = 90] (0,0) -- (1,0) -- (1,1) -- (0,1) -- cycle;
\end{tikzpicture} &
\begin{tikzpicture}
\draw[fill = blue!20, rotate = 180] (0,0) -- (1,0) -- (1,1) -- (0,1) -- cycle;
\end{tikzpicture} \\
\text{C} & \text{D} \\
\begin{tikzpicture}
\draw[fill = blue!20, rotate = 270] (0,0) -- (1,0) -- (1,1) -- (0,1) -- cycle;
\end{tikzpicture} &
\begin{tikzpicture}
\draw[fill = blue!20, rotate = 360] (0,0) -- (1,0) -- (1,1) -- (0,1) -- cycle;
\end{tikzpicture}
\end{array}
\end{document}
"""

with open('figure.tex', 'w') as f:
    f.write(tikz_code)

subprocess.run(['pdflatex', '-interaction=nonstopmode', 'figure.tex'])
os.startfile('figure.pdf')  # Opens PDF (Windows)

from jupyter_tikz import TexFragment

code = r"""\begin{array}{cc} ... \end{array}"""  # Your full array code
tikz = TexFragment(code, tex_packages="tikz,amsmath")
tikz.run_latex(save_image="output.png")  # Saves PNG and displays
