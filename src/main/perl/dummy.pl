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
my $pageLayout = $doc->createPageLayout(
                                         "LandscapeStyle",
                                         properties => {
                                                         'fo:margin-bottom'          => '0.7874in',
                                                         'fo:page-width'             => '11.69"',
                                                         'style:footnote-max-height' => '0in',
                                                         'style:shadow'              => 'none',
                                                         'fo:margin-left'            => '0.7874in',
                                                         'fo:margin-right'           => '0.7874in',
                                                         'fo:page-height'            => '8.27"',
                                                         'style:num-format'          => '1',
                                                         'style:print-orientation'   => 'landscape',
                                                         'style:writing-mode'        => 'lr-tb',
                                                         'fo:margin-top'             => '0.7874in'
                                         }
);
my $page = $doc->createMasterPage(
                                   "LandscapeStyle",
                                   layout => 'landscape',
                                   next   => 'Standard'
);
$doc->switchPageOrientation($page);
$doc->appendParagraph( text  => "Testing",
                       style => "LandscapeStyle" );

$doc->save;
print "\"$docFile\" is saved.\n";
print "Done.";
exit 0;
