# Copy this file to structured-asset-maps.ps1 and edit the Hudu layout IDs/names
# and field labels to match your Hudu instance before running
# . .\jobs\Import-StructuredListAssets.ps1

$HuduStructuredAssetImportMaps = @{
    "Network Devices" = @{
        AssetLayoutId = 31
        NameField     = "Title"

        # SharePoint internal/display field name => Hudu asset field label
        FieldMap = [ordered]@{
            "xogl"  = "Site"
            "Title" = "Device Name"
            "skdq"  = "Type"
            "w5ud"  = "Model"
            "joiu"  = "Serial"
            "zucp"  = "Network"
            "wr9t"  = "IP (s)"
            "qutm"  = "Management Link"
            "h64q"  = "Physical Location"
        }

        PrimaryFieldMap = @{
            PrimaryManufacturer = "_x0069_gy8"
            PrimaryModel        = "w5ud"
            PrimarySerial       = "joiu"
        }
    }

    "ISP Info" = @{
        # AssetLayoutName is resolved with Get-HuduAssetLayouts -Name.
        # Use AssetLayoutId instead if you already know the ID.
        AssetLayoutName = "ISP Info"
        NameField       = "uxyn"

        FieldMap = [ordered]@{
            "uxyn"       = "Name"
            "_x0067_ay1" = "Site Name"
            "saaf"       = "Provider"
            "xufd"       = "Type"
            "voyg"       = "Use"
            "m0mg"       = "Support #"
            "llh6"       = "Hardware Location"
            "fmwe"       = "Hardware Model #"
            "ibmi"       = "Public IP/Subnet"
        }
    }
}
