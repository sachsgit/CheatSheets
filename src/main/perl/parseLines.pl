#!/usr/bin/perl
use strict;
use warnings;
use File::Basename qw(dirname);
use File::Spec::Functions qw(catfile);
use File::Slurper qw(read_lines);

my $rootPath = catfile(dirname($0), "..", "..", "..");
my $mainPath = catfile($rootPath, "src", "main");
my $resourcesPath = catfile($mainPath, "resources");
my $inFile = catfile($resourcesPath, "SQLDeveloper_Shortcuts.txt");

my @lines = read_lines($inFile);

my ($section, $action, $keyStroke) = "" x 3;
my @records;
foreach my $line ( @lines ) {
	if ($line =~ /<value n="context" v="([^"]+)"\/>/) {
		$section = $1;
		($action, $keyStroke) = "" x 2;
	} elsif ($line =~ /<value n="action" v="([^"]+)"\/>/) {
		$action = $1;
		$keyStroke = "";
	} elsif ($line =~ /<value n="key-1" v="([^"]+)"\/>/) {
		$keyStroke = $1;
	}
	
	if ($section ne "" && $action ne "" && $keyStroke ne "") {
		my %rec;
		$rec{'sSect'} = $section;
		$rec{'sAct'} = $action;
		$rec{'sKeys'} = $keyStroke;
		($action, $keyStroke) = "" x 2;
		push @records, \%rec;
	}
} # end foreach line

foreach my $href ( @records ) {
	my %rec = %{$href};
	printf "%60s\t%20s\t%20s\n", $rec{'sSect'}, $rec{'sAct'}, $rec{'sKeys'};
}