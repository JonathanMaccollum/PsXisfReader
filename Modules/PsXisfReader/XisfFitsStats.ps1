class XisfFileStats {
    [decimal]$Exposure
    [string]$Filter
    [string]$Instrument
    [string]$Object
    [decimal]$Gain
    [decimal]$Offset
    [string]$ImageType
    [decimal]$CCDTemp
    [decimal]$SetTemp
    [decimal]$FocalLength
    [decimal]$FocalRatio
    [nullable[DateTime]]$ObsDate
    [nullable[DateTime]]$ObsDateMinus12hr
    [nullable[DateTime]]$LocalDate
    [decimal]$SSWeight
    [decimal]$Pedestal
    [string[]]$History
    [System.IO.FileInfo]$Path
    [decimal]$XPIXSZ
    [decimal]$YPIXSZ
    [string]$Geometry
    [bool]HasTokensInPath([string[]]$tokens){
            $hasToken=$false
            foreach( $x in $tokens) {
                if($this.Path.FullName.ToLower().Contains($x.ToLower())){
                    $hasToken=$true
                    break;
                }
            }
            return $hasToken;
        }
    [bool]IsIntegratedFile() {
        return [bool]($this.History)
    }
}
