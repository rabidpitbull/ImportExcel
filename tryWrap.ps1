$file="c:\temp\testAddress.xlsx"

rm $file -ErrorAction Ignore

function new-testdata {
    param($name,$address)

    [PSCustomObject]@{
        Name=$name
        Address=$address
    }
}

$(
new-testdata "John Doe" "593 Cleveland Street New Britain, `nCT 06051"
new-testdata "John Doe" "8794 Howard Street Elkridge, MD 21075"
new-testdata "John Doe" "9459 Route 20 Jamaica Plain, MA 02130"
new-testdata "John Doe" "8907 9th Street Bayonne, NJ 07002"
new-testdata "John Doe" "6749 Glenwood Avenue Dallas, GA 30132"
new-testdata "John Doe" "8884 Chestnut Avenue Port Saint Lucie, FL 34952"
) | export-excel $file -Show -WrapText -AutoSize