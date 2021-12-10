#!/usr/bin/perl -w
use strict;
use Data::Dumper;
use OpenOffice::OODoc;
my $docFile = "../resources/landscape.odt";

# my $oofile = odfContainer($docFile);
# my $styles = odfDocument(
#                 container => $docFile,
#                 part => 'styles'
# );
# print Dumper($styles);

my $doc = odfDocument( file   => $docFile,
                       create => 'text' );
my $pageLayout = $doc->updatePageLayout(
                        "LandscapeStyle",
                        properties => {
                                        'fo:margin-bottom'          => '0.7874in',
                                        'fo:page-width'             => '11in',
                                        'style:footnote-max-height' => '0in',
                                        'style:shadow'              => 'none',
                                        'fo:margin-left'            => '0.7874in',
                                        'fo:margin-right'           => '0.7874in',
                                        'fo:page-height'            => '8.5in',
                                        'style:num-format'          => '1',
                                        'style:print-orientation'   => 'landscape',
                                        'style:writing-mode'        => 'lr-tb',
                                        'fo:margin-top'             => '0.7874in'
                        }
);
$doc->switchPageOrientation($pageLayout);
$doc->appendParagraph( text  => "Testing",
                       style => $pageLayout );

$doc->save;
print "\"$docFile\" is saved.\n";
print "Done.";
exit 0;
