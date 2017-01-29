Module Module_Well

    Public Class Well
        'Class Well
        Private well_Counter As Integer = 0
        Private str_WellName As String
        Private str_PN As String
        Private str_idwell As String

        'New Well Constructor
        Sub New(ByVal well As String, ByVal pn As String, ByVal idwell As String)
            Me.well_Counter += 1
            Me.str_WellName = well
            Me.str_PN = pn
            Me.str_idwell = idwell
        End Sub

        'TODO:  As ToString method and other needed sub/functions for each well component

    End Class

    Public Class Casing
        'Class Casing 
        Private csg_Counter As Integer = 0
        Private str_CasingName As String
        Private dbl_CasingSetMD As Double
        Private dbl_CasingID As Double
        Private dbl_CasingOD As Double
        Private bln_IsSet As Boolean
        Private bln_IsPlanned As Boolean

        Sub New(ByVal csg_Name As String, ByVal csg_SetMd As Double, ByVal csg_Id As Double, ByVal csg_OD As Double)
            Me.csg_Counter += 1
            Me.str_CasingName = csg_Name
            Me.dbl_CasingSetMD = csg_SetMd
            Me.dbl_CasingID = csg_Id
            Me.dbl_CasingOD = csg_OD
        End Sub

        'TODO:  As ToString method and other needed sub/functions for each well component

    End Class
End Module
