function Get-ADObjectAttributes
{
    [CmdletBinding()]

    param
    (
        [Parameter(Mandatory)]
        [string]
        $ClassName
    )

    Begin
    {
        function Get-RelatedClass
        {
            param
            (
                [Parameter(Mandatory)]
                [string]
                $ClassName
            )

            $Classes = @($ClassName)

            $SubClass = Get-ADObject -SearchBase "$((Get-ADRootDSE).SchemaNamingContext)" -Filter {lDAPDisplayName -eq $ClassName} -properties subClassOf |Select-Object -ExpandProperty subClassOf
            if( $Subclass -and $SubClass -ne $ClassName ) {
                $Classes += Get-RelatedClass $SubClass
            }
            #
            $auxiliaryClasses = Get-ADObject -SearchBase "$((Get-ADRootDSE).SchemaNamingContext)" -Filter {lDAPDisplayName -eq $ClassName} -properties auxiliaryClass | Select-Object -ExpandProperty auxiliaryClass
            foreach( $auxiliaryClass in $auxiliaryClasses ) {
                $Classes += Get-RelatedClass $auxiliaryClass
            }

            $systemAuxiliaryClasses = Get-ADObject -SearchBase "$((Get-ADRootDSE).SchemaNamingContext)" -Filter {lDAPDisplayName -eq $ClassName} -properties systemAuxiliaryClass | Select-Object -ExpandProperty systemAuxiliaryClass
            foreach( $systemAuxiliaryClass in $systemAuxiliaryClasses ) {
                $Classes += Get-RelatedClass $systemAuxiliaryClass
            }
            #
            return $Classes
        }
    }

    Process
    {
        $AllClasses = ( Get-RelatedClass $ClassName | Sort-Object -Unique )

        $AllAttributes = @()

        foreach ($Class in $AllClasses)
        {
            $attributeTypes = 'MayContain','MustContain','systemMayContain','systemMustContain'
            $ClassInfo = Get-ADObject -SearchBase "$((Get-ADRootDSE).SchemaNamingContext)" -Filter {lDAPDisplayName -eq $Class} -properties $attributeTypes
            foreach ($attribute in $attributeTypes)
            {
                $AllAttributes += $ClassInfo.$attribute
            }
        }

        $AllAttributes = ($AllAttributes | Sort-Object -Unique)

        Write-Verbose ('Found {0} attributes for class {1}' -f $AllAttributes.Count, $Class)
        return $AllAttributes
    }
    End
    {
    }
}
