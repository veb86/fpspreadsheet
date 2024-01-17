// include file containing Excel chart hatch patterns
// (used by xlsxooxmlchart.pas)

const
  OOXML_GRAY5_PATTERN =
    'x       '+
    '        '+
    '        '+
    '    x   '+
    '        '+
    '        '+
    '        '+
    '        ';

  OOXML_GRAY10_PATTERN =
    'x       '+
    '        '+
    '    x   '+
    '        '+
    'x       '+
    '        '+
    '    x   '+
    '        ';

  OOXML_GRAY20_PATTERN =
    'x   x   '+
    '        '+
    '  x   x '+
    '        '+
    'x   x   '+
    '        '+
    '  x   x '+
    '        ';

  OOXML_GRAY25_PATTERN =
    'x   x   '+
    '  x   x '+
    'x   x   '+
    '  x   x '+
    'x   x   '+
    '  x   x '+
    'x   x   '+
    '  x   x ';

  OOXML_GRAY30_PATTERN =
    'x x x x '+
    ' x   x  '+
    'x x x x '+
    '   x   x'+
    'x x x x '+
    ' x   x  '+
    'x x x x '+
    '   x   x';

  OOXML_GRAY40_PATTERN =
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x   x'+
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    '   x x x';

  OOXML_GRAY50_PATTERN =
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x';

  OOXML_GRAY60_PATTERN =
    'xxx xxx '+
    ' x x x x'+
    'x xxx xx'+
    ' x x x x'+
    'xxx xxx '+
    ' x x x x'+
    'x xxx xx'+
    ' x x x x';

  OOXML_GRAY70_PATTERN =
    ' xxx xxx'+
    'xx xxx x'+
    ' xxx xxx'+
    'xx xxx x'+
    ' xxx xxx'+
    'xx xxx x'+
    ' xxx xxx'+
    'xx xxx x';

  OOXML_GRAY75_PATTERN =
    ' xxx xxx'+
    'xxxxxxxx'+
    'xx xxx x'+
    'xxxxxxxx'+
    ' xxx xxx'+
    'xxxxxxxx'+
    'xx xxx x'+
    'xxxxxxxx';

  OOXML_GRAY80_PATTERN =
    'xxx xxxx'+
    'xxxxxxxx'+
    'xxxxxxx '+
    'xxxxxxxx'+
    'xxx xxxx'+
    'xxxxxxxx'+
    'xxxxxxx '+
    'xxxxxxxx';

  OOXML_GRAY90_PATTERN =
    ' xxxxxxx'+
    'xxxxxxxx'+
    'xxxxxxxx'+
    'xxxxxxxx'+
    'xxxx xxx'+
    'xxxxxxxx'+
    'xxxxxxxx'+
    'xxxxxxxx';

  OOXML_DASH_DNDIAG_PATTERN =
    'x   x   '+
    ' x   x  '+
    '  x   x '+
    '   x   x'+
    '        '+
    '        '+
    '        '+
    '        ';

  OOXML_DASH_UPDIAG_PATTERN =
    '   x   x'+
    '  x   x '+
    ' x   x  '+
    'x   x   '+
    '        '+
    '        '+
    '        '+
    '        ';

  OOXML_DASH_HORZ_PATTERN =
    'xxxx    '+
    '        '+
    '        '+
    '        '+
    '    xxxx'+
    '        '+
    '        '+
    '        ';

  OOXML_DASH_VERT_PATTERN =
    'x       '+
    'x       '+
    'x       '+
    'x       '+
    '    x   '+
    '    x   '+
    '    x   '+
    '    x   ';

  OOXML_SMALL_CONFETTI_PATTERN =
    'x       '+
    '    x   '+
    ' x      '+
    '      x '+
    '   x    '+
    '       x'+
    '  x     '+
    '     x  ';

  OOXML_LARGE_CONFETTI_PATTERN =
    'x xx   x'+
    '  xx    '+
    '      xx'+
    '   xx xx'+
    'xx xx   '+
    'xx      '+
    '    xx  '+
    'x   xx x';

  OOXML_ZIGZAG_PATTERN =
    'x      x'+
    ' x    x '+
    '  x  x  '+
    '   xx   '+
    'x      x'+
    ' x    x '+
    '  x  x  '+
    '   xx   ';

  OOXML_WAVE_PATTERN =
    '        '+
    '   xx   '+
    '  x  x x'+
    'xx      '+
    '        '+
    '   xx   '+
    '  x  x x'+
    'xx      ';

  OOXML_DIAG_BRICK_PATTERN =
    '       x'+
    '      x '+
    '     x  '+
    '    x   '+
    '   xx   '+
    '  x  x  '+
    ' x    x '+
    'x      x';

  OOXML_HORZ_BRICK_PATTERN =
    'xxxxxxxx'+
    'x       '+
    'x       '+
    'x       '+
    'xxxxxxxx'+
    '    x   '+
    '    x   '+
    '    x   ';

  OOXML_WEAVE_PATTERN =
    'x   x   '+
    ' x x x  '+
    '  x   x '+
    ' x   x x'+
    'x   x   '+
    '   x x  '+
    '  x   x '+
    ' x x   x';

  OOXML_PLAID_PATTERN =
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x'+
    'xxxx    '+
    'xxxx    '+
    'xxxx    '+
    'xxxx    ';

  OOXML_DIVOT_PATTERN =
    'x       '+
    ' x      '+
    'x       '+
    '        '+
    '     x  '+
    '    x   '+
    '     x  '+
    '        ';

  OOXML_DOT_GRID_PATTERN =
    'x x x x '+
    '        '+
    'x       '+
    '        '+
    'x       '+
    '        '+
    'x       '+
    '        ';

  OOXML_DOT_DIAMOND_PATTERN =
    'x       '+
    '        '+
    '  x   x '+
    '        '+
    '    x   '+
    '        '+
    '  x   x '+
    '        ';

  OOXML_SHINGLE_PATTERN =
    '      xx'+
    'x    x  '+
    ' x  x   '+
    '  xx    '+
    '    xx  '+
    '      x '+
    '       x'+
    '       x';

  OOXML_TRELLIS_PATTERN =
    'xxxxxxxx'+
    ' xx  xx '+
    'xxxxxxxx'+
    'x  xx  x'+
    'xxxxxxxx'+
    ' xx  xx '+
    'xxxxxxxx'+
    'x  xx  x';

  OOXML_SPHERE_PATTERN =
    ' xxx xxx'+
    'x  xx   '+
    'xxxxx   '+
    'xxxxx   '+
    ' xxx xxx'+
    'x   x  x'+
    'x   xxxx'+
    'x   xxxx';

  OOXML_SMALL_CHECKERBOARD_PATTERN =
    'xx  xx  '+
    'xx  xx  '+
    '  xx  xx'+
    '  xx  xx'+
    'xx  xx  '+
    'xx  xx  '+
    '  xx  xx'+
    '  xx  xx';

  OOXML_LARGE_CHECKBOARD_PATTERN =
    'xxxx    '+
    'xxxx    '+
    'xxxx    '+
    'xxxx    '+
    '    xxxx'+
    '    xxxx'+
    '    xxxx'+
    '    xxxx';

  OOXML_SOLID_DIAMOND_PATTERN =
    '   x    '+
    '  xxx   '+
    ' xxxxx  '+
    'xxxxxxx '+
    ' xxxxx  '+
    '  xxx   '+
    '   x    '+
    '        ';

