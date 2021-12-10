#!/usr/bin/perl -w
use strict;
use constant::boolean;
use Data::Dump qw(dump);
use OpenOffice::OODoc;
use Text::CSV;
use Text::Trim;

use constant DEBUG => FALSE;

my $inFile  = "../resources/CSV Files/openscript-eclipse-3.4.csv";
my $docFile = "../resources/openscript-eclipse-3.4.odt";

my @records;
my @sections;

my $csv = Text::CSV->new(
                          {
                            binary    => 1,
                            auto_diag => 1
                          }
);

open my $fh, "<", $inFile or die "$inFile: $!";
while ( my $row = $csv->getline($fh) ) {
    my @fields = @$row;
    dump(@fields) if ( DEBUG == TRUE );
    my $record = {
                   SECTION     => trim( $fields[0] ),
                   DESCRIPTION => trim( $fields[1] ),
                   KEYS        => trim( $fields[2] ),
                   AREA        => trim( $fields[3] )
    };
    push @records,  $record;
    push @sections, $record->{SECTION} unless grep { $_ eq $record->{SECTION} } @sections;
}
close $fh;

unlink($docFile) if ( -f $docFile );
my $doc = odfDocument( file   => $docFile,
                       create => 'text' );
my $entryStyle = $doc->createStyle(
    "EntryStyle",
    family     => 'paragraph',
    parent     => 'Standard',
    properties => {
                    '-area'          => 'text',
                    'fo:font-size'   => '8pt',
                    'fo:color'       => '#88888',         # Grey
                    'fo:font-family' => 'Arial Narrow',
    }
);
my $headerStyle = $doc->createStyle(
    "HeaderStyle",
    family     => 'paragraph',
    parent     => 'Standard',
    properties => {
                    -area            => 'text',
                    'fo:font-size'   => '11pt',
                    'fo:font-weight' => 'bold',
                    'fo:color'       => '#000000',        # black
                    'fo:font-family' => 'Arial',
    }
);
$doc->createStyle(
                   "border",
                   family     => 'table-cell',
                   parent     => 'Standard',
                   properties => {
                                   'fo:border'  => "0.0cm solid #000000",
                                   'fo:padding' => "0.0cm",
                   }
);
my $sectionCounter = 1;
foreach my $sec (@sections) {
    $doc->appendParagraph( text  => $sec,
                           style => "HeaderStyle" );
    my @line;
    my $lineCounter = 0;
    my $tableName   = "table" . $sectionCounter;
    foreach my $rec (@records) {
        if ( $rec->{SECTION} eq $sec ) {
            $line[$lineCounter] = ( [ $rec->{DESCRIPTION}, $rec->{KEYS}, $rec->{AREA} ] );
            $lineCounter++;
            printf "%-25s\t%-15s\t%-15s\r\n", $rec->{DESCRIPTION}, $rec->{KEYS}, $rec->{AREA}
              if ( DEBUG == TRUE );
        }    # End If Correct Section
    }    # End For Each Record

    my $table = $doc->appendTable( $tableName, $lineCounter, 3 );
    for ( my $r = 0 ; $r < $lineCounter ; $r++ ) {
        for ( my $c = 0 ; $c < 3 ; $c++ ) {
            print "\$line[$r][$c] = $line[$r][$c]\n" if ( DEBUG == TRUE );
            $doc->cellStyle( $tableName, $r, $c, 'border' );
            $doc->cellValue( $tableName, $r, $c, $line[$r][$c] );
            $doc->setAttributes( $doc->getCellParagraph( $tableName, $r, $c ),
                                 'text:style-name' => 'EntryStyle' );
        }
    }    # End For Each Line
    $sectionCounter++;
}

$doc->save;
print "\"$docFile\" is saved.\n";
print "Done.";
exit 0;
