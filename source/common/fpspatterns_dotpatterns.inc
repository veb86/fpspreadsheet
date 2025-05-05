// Include file containing built-in chart dot patterns
// (used by fpspatterns.pas)

const
  GRAY06_PATTERN =
    'x       '+
    '        '+
    '    x   '+
    '        '+
    'x       '+
    '        '+
    '    x   '+
    '        ';

  GRAY12_PATTERN =
    'x   x   '+
    '        '+
    '  x   x '+
    '        '+
    'x   x   '+
    '        '+
    '  x   x '+
    '        ';

  GRAY20_PATTERN =
    'x   x   '+
    '        '+
    '  x   x '+
    '        '+
    'x   x   '+
    '        '+
    '  x   x '+
    '        ';

  GRAY25_PATTERN =
    'x   x   '+
    '  x   x '+
    'x   x   '+
    '  x   x '+
    'x   x   '+
    '  x   x '+
    'x   x   '+
    '  x   x ';

  GRAY30_PATTERN =
    'x x x x '+
    ' x   x  '+
    'x x x x '+
    '   x   x'+
    'x x x x '+
    ' x   x  '+
    'x x x x '+
    '   x   x';

  GRAY40_PATTERN =
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x   x'+
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    '   x x x';

  GRAY50_PATTERN =
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x';

  GRAY60_PATTERN =
    'xxx xxx '+
    ' x x x x'+
    'x xxx xx'+
    ' x x x x'+
    'xxx xxx '+
    ' x x x x'+
    'x xxx xx'+
    ' x x x x';

  GRAY70_PATTERN =
    ' xxx xxx'+
    'xx xxx x'+
    ' xxx xxx'+
    'xx xxx x'+
    ' xxx xxx'+
    'xx xxx x'+
    ' xxx xxx'+
    'xx xxx x';

  GRAY75_PATTERN =
    ' xxx xxx'+
    'xxxxxxxx'+
    'xx xxx x'+
    'xxxxxxxx'+
    ' xxx xxx'+
    'xxxxxxxx'+
    'xx xxx x'+
    'xxxxxxxx';

  GRAY80_PATTERN =
    'xxx xxxx'+
    'xxxxxxxx'+
    'xxxxxxx '+
    'xxxxxxxx'+
    'xxx xxxx'+
    'xxxxxxxx'+
    'xxxxxxx '+
    'xxxxxxxx';

  GRAY90_PATTERN =
    ' xxxxxxx'+
    'xxxxxxxx'+
    'xxxxxxxx'+
    'xxxxxxxx'+
    'xxxx xxx'+
    'xxxxxxxx'+
    'xxxxxxxx'+
    'xxxxxxxx';

  HOR_PATTERN_THIN =
    'xxxxxxxx'+
    '        '+
    '        '+
    '        '+
    '        '+
    '        '+
    '        '+
    '        ';

  HOR_PATTERN_THICK =
    'xxxxxxxx'+
    'xxxxxxxx'+
    '        '+
    '        '+
    '        '+
    '        '+
    '        '+
    '        ';

  HOR_PATTERN_NARROW =
    'xxxxxxxx'+
    '        '+
    '        '+
    '        '+
    'xxxxxxxx'+
    '        '+
    '        '+
    '        ';

  VERT_PATTERN_THIN =
    'x       '+
    'x       '+
    'x       '+
    'x       '+
    'x       '+
    'x       '+
    'x       '+
    'x       ';

  VERT_PATTERN_THICK =
    'xx      '+
    'xx      '+
    'xx      '+
    'xx      '+
    'xx      '+
    'xx      '+
    'xx      '+
    'xx      ';

  VERT_PATTERN_NARROW =
    'x   x   '+
    'x   x   '+
    'x   x   '+
    'x   x   '+
    'x   x   '+
    'x   x   '+
    'x   x   '+
    'x   x   ';

  DIAG_DOWN_PATTERN_THIN =
    'x       '+
    ' x      '+
    '  x     '+
    '   x    '+
    '    x   '+
    '     x  '+
    '      x '+
    '       x';

  DIAG_DOWN_PATTERN_THICK =
    'xx      '+
    ' xx     '+
    '  xx    '+
    '   xx   '+
    '    xx  '+
    '     xx '+
    '      xx'+
    'x      x';

  DIAG_DOWN_PATTERN_NARROW =
    'x   x   '+
    ' x   x  '+
    '  x   x '+
    '   x   x'+
    'x   x   '+
    ' x   x  '+
    '  x   x '+
    '   x   x';

  DIAG_UP_PATTERN_THIN =
    '       x'+
    '      x '+
    '     x  '+
    '    x   '+
    '   x    '+
    '  x     '+
    ' x      '+
    'x       ';

  DIAG_UP_PATTERN_THICK =
    'x      x'+
    '      xx'+
    '     xx '+
    '    xx  '+
    '   xx   '+
    '  xx    '+
    ' xx     '+
    'xx      ';

  DIAG_UP_PATTERN_NARROW =
    '   x   x'+
    '  x   x '+
    ' x   x  '+
    'x   x   '+
    '   x   x'+
    '  x   x '+
    ' x   x  '+
    'x   x   ';

  HATCH_PATTERN_THIN =
    '       x'+
    'x     x '+
    ' x   x  '+
    '  x x   '+
    '   x    '+
    '  x x   '+
    ' x   x  '+
    'x     x ';

  HATCH_PATTERN_THICK =
    'x      x'+
    'xx    xx'+
    ' xx  xx '+
    '  xxxx  '+
    '   xx   '+
    '  xxxx  '+
    ' xx  xx '+
    'xx    xx';

  HATCH_PATTERN_NARROW =
    '   x   x'+
    'x x x x '+
    ' x   x  '+
    'x x x x '+
    '   x   x'+
    'x x x  x'+
    ' x   x  '+
    'x x x x ';

  HATCH_PATTERN_DOT =
    'x       '+
    '        '+
    '  x   x '+
    '        '+
    '    x   '+
    '        '+
    '  x   x '+
    '        ';

  CROSS_PATTERN_THIN =
    '   x    '+
    '   x    '+
    '   x    '+
    'xxxxxxxx'+
    '   x    '+
    '   x    '+
    '   x    '+
    '   x    ';

  CROSS_PATTERN_THICK =
    '   xx   '+
    '   xx   '+
    '   xx   '+
    'xxxxxxxx'+
    'xxxxxxxx'+
    '   xx   '+
    '   xx   '+
    '   xx   ';

  CROSS_PATTERN_NARROW =
    '  x   x '+
    'xxxxxxxx'+
    '  x   x '+
    '  x   x '+
    '  x   x '+
    'xxxxxxxx'+
    '  x   x '+
    '  x   x ';

  CROSS_PATTERN_DOT =
    'x x x x '+
    '        '+
    'x       '+
    '        '+
    'x       '+
    '        '+
    'x       '+
    '        ';

  DIAG_DOWN_DASH_PATTERN =
    'x   x   '+
    ' x   x  '+
    '  x   x '+
    '   x   x'+
    '        '+
    '        '+
    '        '+
    '        ';

  DIAG_UP_DASH_PATTERN =
    '   x   x'+
    '  x   x '+
    ' x   x  '+
    'x   x   '+
    '        '+
    '        '+
    '        '+
    '        ';

  HOR_DASH_PATTERN =
    'xxxx    '+
    '        '+
    '        '+
    '        '+
    '    xxxx'+
    '        '+
    '        '+
    '        ';

  VERT_DASH_PATTERN =
    'x       '+
    'x       '+
    'x       '+
    'x       '+
    '    x   '+
    '    x   '+
    '    x   '+
    '    x   ';

  CHECKERBOARD_LARGE_PATTERN =
    'xxxx    '+
    'xxxx    '+
    'xxxx    '+
    'xxxx    '+
    '    xxxx'+
    '    xxxx'+
    '    xxxx'+
    '    xxxx';

  CHECKERBOARD_SMALL_PATTERN =
    'xx  xx  '+
    'xx  xx  '+
    '  xx  xx'+
    '  xx  xx'+
    'xx  xx  '+
    'xx  xx  '+
    '  xx  xx'+
    '  xx  xx';

  CONFETTI_LARGE_PATTERN =
    'x xx   x'+
    '  xx    '+
    '      xx'+
    '   xx xx'+
    'xx xx   '+
    'xx      '+
    '    xx  '+
    'x   xx x';

  CONFETTI_SMALL_PATTERN =
    'x       '+
    '    x   '+
    ' x      '+
    '      x '+
    '   x    '+
    '       x'+
    '  x     '+
    '     x  ';

  DIAMOND_PATTERN =
    '   x    '+
    '  xxx   '+
    ' xxxxx  '+
    'xxxxxxx '+
    ' xxxxx  '+
    '  xxx   '+
    '   x    '+
    '        ';

  BRICK_HOR_PATTERN =
    'xxxxxxxx'+
    'x       '+
    'x       '+
    'x       '+
    'xxxxxxxx'+
    '    x   '+
    '    x   '+
    '    x   ';

  BRICK_DIAG_PATTERN =
    '       x'+
    '      x '+
    '     x  '+
    '    x   '+
    '   xx   '+
    '  x  x  '+
    ' x    x '+
    'x      x';

  DIVOT_PATTERN =
    'x       '+
    ' x      '+
    'x       '+
    '        '+
    '     x  '+
    '    x   '+
    '     x  '+
    '        ';

  PLAID_PATTERN =
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x'+
    'xxxx    '+
    'xxxx    '+
    'xxxx    '+
    'xxxx    ';

  SHINGLE_PATTERN =
    '      xx'+
    'x    x  '+
    ' x  x   '+
    '  xx    '+
    '    xx  '+
    '      x '+
    '       x'+
    '       x';

  SPHERE_PATTERN =
    ' xxx xxx'+
    'x  xx   '+
    'xxxxx   '+
    'xxxxx   '+
    ' xxx xxx'+
    'x   x  x'+
    'x   xxxx'+
    'x   xxxx';

  TRELLIS_PATTERN =
    'xxxxxxxx'+
    ' xx  xx '+
    'xxxxxxxx'+
    'x  xx  x'+
    'xxxxxxxx'+
    ' xx  xx '+
    'xxxxxxxx'+
    'x  xx  x';

  WAVE_PATTERN =
    '        '+
    '   xx   '+
    '  x  x x'+
    'xx      '+
    '        '+
    '   xx   '+
    '  x  x x'+
    'xx      ';

  WEAVE_PATTERN =
    'x   x   '+
    ' x x x  '+
    '  x   x '+
    ' x   x x'+
    'x   x   '+
    '   x x  '+
    '  x   x '+
    ' x x   x';

  ZIGZAG_PATTERN =
    'x      x'+
    ' x    x '+
    '  x  x  '+
    '   xx   '+
    'x      x'+
    ' x    x '+
    '  x  x  '+
    '   xx   ';

