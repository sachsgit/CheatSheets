#!/usr/bin/perl
use strict;
use OpenOffice::OODoc;

# URL: https://www.perlmonks.org/?part=1;displaytype=displaycode;abspart=1;node_id=1166574

my $doc = odfDocument( file => 'table_test.odt', create => 'text' );

$doc->createStyle(
                   "border",
                   family     => 'table-cell',
                   parent     => 'Standard',
                   properties => {
                                   'fo:border'  => "0.05cm solid #000000",
                                   'fo:padding' => "0.05cm",
                   }
);

$doc->createStyle(
                   "2cm",
                   family     => 'table-column',
                   parent     => 'Standard',
                   properties => { 'style:column-width' => "2cm" }
);

$doc->createStyle(
                   "6cm",
                   family     => 'table-column',
                   parent     => 'Standard',
                   properties => { 'style:column-width' => "6cm" }
);

$doc->createStyle(
    "theading",
    family     => 'paragraph',
    parent     => 'Standard',
    properties => {
                    -area            => 'text',
                    'fo:font-size'   => '13pt',
                    'fo:font-weight' => 'bold',
                    'fo:color'       => '#000000',    # black
                    'fo:font-family' => 'arial',
    }
);
$doc->updateStyle(
    "theading",
    properties => {
                    'fo:background-color' => "#c0c0c0",    # light grey
    }
);

my @width = ( '2cm', '6cm', '2cm' );
my @data = (
             [ 'S No', 'Name', 'Value' ],
             [ 'A1',   'B1',   'C1' ],
             [ 'A2',   'B2',   'C2' ],
             [ 'A3',   'B3',   'C3' ],
             [ 'A4',   'B4',   'C4' ]
);

my $table = $doc->appendTable( 'tab1', 5, 3 );

for my $c ( 0 .. 2 ) {
    $doc->columnStyle( 'tab1', $c, $width[$c] );
    for my $r ( 0 .. 4 ) {
        $doc->cellStyle( 'tab1', $r, $c, 'border' );
        $doc->cellValue( 'tab1', $r, $c, $data[$r][$c] );
    }
    $doc->setAttributes( $doc->getCellParagraph( "tab1", 0, $c ), 'text:style-name' => 'theading' );
}
foreach my $fd ( $doc->getFontDeclarations ) {
    print $doc->getFontName($fd) . "\n";
}
$doc->save;
