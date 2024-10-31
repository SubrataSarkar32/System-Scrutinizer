# -*- coding: utf-8 -*-
__version__ = '1.0.2 (beta)'

red = "\u001b[38;5;160m"
light_red = "\u001b[38;5;208m"
green = "\033[0;32m"
yellow = "\u001b[38;5;184m"
nocolor = "\033[0m"
cyan = "\033[36m"
ascii_col = "\033[0;92m"
log_col = "\u001b[38;5;110m"

passFlag = 'src="./assets/green.png" alt="pass"'
failFlag = 'src="./assets/red.png" alt="fail"'
whiteFlag = 'src="./assets/white.png" alt="partially pass"'

strikethrough = ' style="text-decoration: line-through;"'

referenceLink = 'javascript:void(0)'

debug_mode = True

if __name__ == '__main__':
    print("This module cannot be run directly. Please import it.")