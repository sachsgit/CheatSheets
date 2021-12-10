use Text::Table;
my $tb = Text::Table->new(
    #"Planet", "Radius\nkm", "Density\ng/cm^3"
);
$tb->load(
    [ "Mercury", 2360, 3.7 ],
    [ "Venus", 6110, 5.1 ],
    [ "Earth", 6378, 5.52 ],
    [ "Jupiter", 71030, 1.3 ],
);
print $tb;