Attribute VB_Name = "NZS1170_Seismic_Coefficient"
'Refer https://engineervsheep.com/2020/seismic-coefficient-1/ for further notes on usage
'____________________________________________________________________________________________________________
Option Explicit

Function Loading_ADRS(T_1 As Variant, Site_subsoil_class As String, Hazard_factor As Double, Return_period_factor As Double, _
                      Fault_distance As Variant, S_p As Double, zeta_sys As Double, _
                      Optional D_subsoil_interpolate As Boolean = False, Optional T_site As Double = 1.5) As Variant
'function to calculate the spectral displacement and spectral acceleration for plotting an ADRS Curve (Acceleration Displacement Response Spectrum)

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_ADRS(T_1,Site_subsoil_class,Hazard_factor,Return_period_factor,Fault_distance,S_p,zeta_sys,D_subsoil_interpolate,T_site)
'T_1 = FIRST MODE PERIOD
'Site_subsoil_class = SITE SUBSOIL CLASS, i.e. A/B/C/D/E (ENTERED AS A STRING)
'Hazard_factor = HAZARD FACTOR Z
'Return_period_factor = RETURN PERIOD FACTOR Ru OR Rs
'Fault_distance = THE SHORTEST DISTANCE (IN kM's) FROM THE SITE TO THE NEAREST FAULT LISTED IN TABLE 3.6 OF NZS1170.5
'                 IF NOT RELEVANT USE "N/A" OR A NUMBER >= 20
'S_p = STRUCTURAL PERFORMANCE FACTOR
'zeta_sys = SYSTEM DAMPING (AS A PERCENTAGE)
'OPTIONAL - D_subsoil_interpolation is optional - TRUE/FALSE - consider interpolation for class D soils
'OPTIONAL - T_site is optional - Site period when required for considering interpolation for class D soils
'           (note default value of 1.5 seconds equates to no interpolation)
'________________________________________________________________________________________________________________

    Dim results As Variant
    Dim C_d_T As Variant
    Dim K_zeta As Double
    Dim i

    'determine the spectral damping reduction factor
    K_zeta = Loading_K_zeta(zeta_sys)

    'determine the 5% damping Response Spectrum design spectrum with ductility of 1.0 (so k_mu = 1.0)
    C_d_T = Loading_C_d_T(T_1, Site_subsoil_class, Hazard_factor, Return_period_factor, Fault_distance, 1, S_p, False, D_subsoil_interpolate, T_site)

    ReDim results(LBound(T_1) To UBound(T_1), 1 To 2)

    For i = LBound(T_1) To UBound(T_1)

        'determine the spectral acceleration (S_a)
        results(i, 2) = K_zeta * C_d_T(i, 1)
        'determine the spectral displacement (S_d)
        results(i, 1) = Loading_K_delta_T(T_1(i, 1)) * results(i, 2)

    Next i

    'return results
    Loading_ADRS = results

End Function

Function Loading_K_delta_T(T_1 As Variant) As Variant
'function to calculate the Displacement spectral scaling factor

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_K_delta_T(T_1)
'T_1 = FIRST MODE PERIOD
'________________________________________________________________________________________________________________

    Dim pi As Double
    pi = 3.14159265358979

    Loading_K_delta_T = 9810 * T_1 ^ 2 / (4 * pi ^ 2)

End Function

Function Loading_K_zeta(zeta_sys As Double)
'function to calculate the spectral damping reduction factor

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_K_zeta(zeta_sys)
'zeta_sys = SYSTEM DAMPING (AS A PERCENTAGE)
'________________________________________________________________________________________________________________

    Loading_K_zeta = (7 / (2 + zeta_sys)) ^ 0.5

End Function

Function Loading_C_h_T(T_1 As Double, Site_subsoil_class As String, Optional ESM_case As Boolean = True, _
                       Optional D_subsoil_interpolate As Boolean = False, Optional T_site As Double = 1.5)
'Function to calculate the spectral shape factor (Ch(T)) based on site subsoil soil class (A/B/C/D/E) type
'and first mode period

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_C_h_T(T_1,Site_subsoil_class,ESM_case,D_subsoil_interpolate,T_site)
'T_1 = FIRST MODE PERIOD
'Site_subsoil_class = SITE SUBSOIL CLASS, i.e. A/B/C/D/E (ENTERED AS A STRING)
'OPTIONAL - ESM_case is optional (ESM = Equivalent Static Method), when set to false the Response Spectrum Method (RSM)
'           shape factor will be calculated.
'OPTIONAL - D_subsoil_interpolation is optional - TRUE/FALSE - consider interpolation for class D soils
'OPTIONAL - T_site is optional - Site period when required for considering interpolation for class D soils
'           (note default value of 1.5 seconds equates to no interpolation)
'________________________________________________________________________________________________________________

'set entry to uppercase
    Site_subsoil_class = UCase(Site_subsoil_class)

    Dim C_h_T_shallow

    'test for ESM vs RSM spectrum analysis
    If ESM_case = True Then
        'ESM case
        'case based on soil type

        Select Case Site_subsoil_class
                'values/formulas taken from the NZS1170.5 commentary
            Case "A", "B"

                If T_1 >= 0 And T_1 < 0.4 Then Loading_C_h_T = 1.89
                If T_1 >= 0.4 And T_1 <= 1.5 Then Loading_C_h_T = 1.6 * (0.5 / T_1) ^ 0.75
                If T_1 > 1.5 And T_1 <= 3 Then Loading_C_h_T = 1.05 / T_1
                If T_1 > 3 Then Loading_C_h_T = 3.15 / T_1 ^ 2

                'additional check because at just above T=0.4 the value spikes above 1.89, this smooths out this blip
                If Loading_C_h_T >= 1.89 Then Loading_C_h_T = 1.89

            Case "C"

                If T_1 >= 0 And T_1 < 0.4 Then Loading_C_h_T = 2.36
                If T_1 >= 0.4 And T_1 <= 1.5 Then Loading_C_h_T = 2 * (0.5 / T_1) ^ 0.75
                If T_1 > 1.5 And T_1 <= 3 Then Loading_C_h_T = 1.32 / T_1
                If T_1 > 3 Then Loading_C_h_T = 3.96 / T_1 ^ 2

                'additional check because at just above T=0.4 the value spikes above 2.36, this smooths out this blip
                If Loading_C_h_T > 2.36 Then Loading_C_h_T = 2.36

            Case "D"

                'check if interpolation can be adopted
                If D_subsoil_interpolate = True And T_site >= 0.6 And T_site <= 1.5 And T_1 >= 0.1 Then
                    'interpolate between C & D site subsoil classes considered

                    'min T=0.4 seconds for ESM case with interpolation
                    If T_1 < 0.4 Then T_1 = 0.4

                    'Class C values for NITH case
                    If T_1 >= 0.4 And T_1 <= 1.5 Then C_h_T_shallow = 2 * (0.5 / T_1) ^ 0.75
                    If T_1 > 1.5 And T_1 <= 3 Then C_h_T_shallow = 1.32 / T_1
                    If T_1 > 3 Then C_h_T_shallow = 3.96 / T_1 ^ 2

                    Loading_C_h_T = C_h_T_shallow * (1 + 0.5 * (T_site - 0.25))

                    'additional check because at just above T=0.4 the value spikes above 3, this smooths out this blip
                    If Loading_C_h_T > 3 Then Loading_C_h_T = 3

                Else
                    'No interpolate considered, or able to be considered
                    If T_1 >= 0 And T_1 < 0.56 Then Loading_C_h_T = 3
                    If T_1 >= 0.56 And T_1 <= 1.5 Then Loading_C_h_T = 2.4 * (0.75 / T_1) ^ 0.75
                    If T_1 > 1.5 And T_1 <= 3 Then Loading_C_h_T = 2.14 / T_1
                    If T_1 > 3 Then Loading_C_h_T = 6.42 / T_1 ^ 2

                    'additional check because at just below T=0.56 the value steps abruptly, this smooths out this blip
                    If Loading_C_h_T > 3 Then Loading_C_h_T = 3

                End If

            Case "E"

                If T_1 >= 0 And T_1 < 1 Then Loading_C_h_T = 3
                If T_1 >= 1 And T_1 <= 1.5 Then Loading_C_h_T = 3 / T_1 ^ 0.75
                If T_1 > 1.5 And T_1 <= 3 Then Loading_C_h_T = 3.32 / T_1
                If T_1 > 3 Then Loading_C_h_T = 9.96 / T_1 ^ 2

            Case Else
                'error?

        End Select

    Else
        'RSM case
        'case based on soil type

        Select Case Site_subsoil_class
                'values/formulas taken from the NZS1170.5 commentary
            Case "A", "B"

                If T_1 >= 0 And T_1 < 0.1 Then Loading_C_h_T = 1 + 1.35 * (T_1 / 0.1)
                If T_1 >= 0.1 And T_1 < 0.3 Then Loading_C_h_T = 2.35
                If T_1 >= 0.3 And T_1 <= 1.5 Then Loading_C_h_T = 1.6 * (0.5 / T_1) ^ 0.75
                If T_1 > 1.5 And T_1 <= 3 Then Loading_C_h_T = 1.05 / T_1
                If T_1 > 3 Then Loading_C_h_T = 3.15 / T_1 ^ 2

            Case "C"

                If T_1 >= 0 And T_1 < 0.1 Then Loading_C_h_T = 1.33 + 1.6 * (T_1 / 0.1)
                If T_1 >= 0.1 And T_1 < 0.3 Then Loading_C_h_T = 2.93
                If T_1 >= 0.3 And T_1 <= 1.5 Then Loading_C_h_T = 2 * (0.5 / T_1) ^ 0.75
                If T_1 > 1.5 And T_1 <= 3 Then Loading_C_h_T = 1.32 / T_1
                If T_1 > 3 Then Loading_C_h_T = 3.96 / T_1 ^ 2

                'additional check because at just above T=0.3 the value spikes above 2.93, this smooths out this blip
                If Loading_C_h_T > 2.93 Then Loading_C_h_T = 2.93

            Case "D"

                'check if interpolation can be adopted
                If D_subsoil_interpolate = True And T_site >= 0.6 And T_site <= 1.5 And T_1 >= 0.1 Then
                    'interpolate between C & D site subsoil classes considered

                    'Class C values for NITH case
                    If T_1 >= 0.1 And T_1 < 0.3 Then C_h_T_shallow = 2.93
                    If T_1 >= 0.3 And T_1 <= 1.5 Then C_h_T_shallow = 2 * (0.5 / T_1) ^ 0.75
                    If T_1 > 1.5 And T_1 <= 3 Then C_h_T_shallow = 1.32 / T_1
                    If T_1 > 3 Then C_h_T_shallow = 3.96 / T_1 ^ 2

                    'additional check because at just above T=0.3 the value spikes above 2.93, this smooths out this blip
                    If C_h_T_shallow > 2.93 Then C_h_T_shallow = 2.93

                    Loading_C_h_T = C_h_T_shallow * (1 + 0.5 * (T_site - 0.25))
                    If Loading_C_h_T > 3 Then Loading_C_h_T = 3

                Else
                    'No interpolate considered, or able to be considered
                    If T_1 >= 0 And T_1 < 0.1 Then Loading_C_h_T = 1.12 + 1.88 * (T_1 / 0.1)
                    If T_1 >= 0.1 And T_1 < 0.56 Then Loading_C_h_T = 3
                    If T_1 >= 0.56 And T_1 <= 1.5 Then Loading_C_h_T = 2.4 * (0.75 / T_1) ^ 0.75
                    If T_1 > 1.5 And T_1 <= 3 Then Loading_C_h_T = 2.14 / T_1
                    If T_1 > 3 Then Loading_C_h_T = 6.42 / T_1 ^ 2

                End If

            Case "E"

                If T_1 >= 0 And T_1 < 0.1 Then Loading_C_h_T = 1.12 + 1.88 * (T_1 / 0.1)
                If T_1 >= 0.1 And T_1 < 1 Then Loading_C_h_T = 3
                If T_1 >= 1 And T_1 <= 1.5 Then Loading_C_h_T = 3 / T_1 ^ 0.75
                If T_1 > 1.5 And T_1 <= 3 Then Loading_C_h_T = 3.32 / T_1
                If T_1 > 3 Then Loading_C_h_T = 9.96 / T_1 ^ 2

            Case Else
                'error?

        End Select

    End If

End Function

Function Loading_k_mu(T_1 As Double, Site_subsoil_class As String, mu As Double)
'Function to calculate the inelastic spectrum scaling factor

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_k_mu(T_1,Site_subsoil_class,mu)
'T_1 = FIRST MODE PERIOD
'Site_subsoil_class = SITE SUBSOIL CLASS, i.e. A/B/C/D/E (ENTERED AS A STRING)
'mu = DUCTILITY
'________________________________________________________________________________________________________________

'set entry to uppercase
    Site_subsoil_class = UCase(Site_subsoil_class)

    'calculate based on minimum of T_1 = 0.4 as per NZS110.5 CL5.2.1.1
    If T_1 < 0.4 Then T_1 = 0.4

    Select Case Site_subsoil_class

        Case "A", "B", "C", "D"

            If T_1 >= 0.7 Then
                Loading_k_mu = mu
            Else
                Loading_k_mu = (mu - 1) * T_1 / 0.7 + 1
            End If

        Case "E"

            If T_1 >= 1 Or mu < 1.5 Then
                Loading_k_mu = mu
            Else
                Loading_k_mu = (mu - 1.5) * T_1 + 1.5
            End If

    End Select

End Function

Function Loading_N_max_T(T_1 As Double)
'Function to calculate Nmax(T), the maximum near fault factor

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_N_max_T(T_1)
'T_1 = FIRST MODE PERIOD
'________________________________________________________________________________________________________________

    If T_1 <= 1.5 Then Loading_N_max_T = 1
    If T_1 >= 5 Then Loading_N_max_T = 1.72
    If T_1 > 1.5 And T_1 <= 4 Then Loading_N_max_T = 0.24 * T_1 + 0.64
    If T_1 > 4 And T_1 < 5 Then Loading_N_max_T = 0.12 * T_1 + 1.12

End Function

Function Loading_N_T_D(T_1 As Double, Fault_distance As Variant, Return_period_factor As Double)
'Function to calculate N(T,D), the near fault factor
'calculation is based on R_u or R_s (Return Period Factor) rather than the probability of exceedance as its
'easier to follow when written like this

'N(T,D) = 1.0 for R <= 0.75
'N(T,D) = varies for R > 0.75

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_N_T_D(T_1,Fault_distance,Return_period_factor)
'T_1 = FIRST MODE PERIOD
'Fault_distance = THE SHORTEST DISTANCE (IN kM's) FROM THE SITE TO THE NEAREST FAULT LISTED IN TABLE 3.6 OF NZS1170.5
'                 IF NOT RELEVANT USE "N/A" OR A NUMBER >= 20
'Return_period_factor = RETURN PERIOD FACTOR Ru OR Rs
'________________________________________________________________________________________________________________

    If Return_period_factor <= 0.75 Then Loading_N_T_D = 1

    If Return_period_factor > 0.75 Then

        If Fault_distance <= 2 Then Loading_N_T_D = Loading_N_max_T(T_1)
        If Fault_distance = "N/A" Or Fault_distance > 20 Then Loading_N_T_D = 1
        If Fault_distance > 2 And Fault_distance <= 20 Then Loading_N_T_D = 1 + (Loading_N_max_T(T_1) - 1) * (20 - Fault_distance) / 18

    End If

End Function

Private Function Loading_D_subsoil_interpolate_test(T_1 As Double, Site_subsoil_class As String, T_site As Double)
'function to output whether or not interpolation can be used for D type soils (interpolating between C and D)

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_D_subsoil_interpolate_test(T_1,Site_subsoil_class,T_site)
'T_1 = FIRST MODE PERIOD
'Site_subsoil_class = SITE SUBSOIL CLASS, i.e. A/B/C/D/E (ENTERED AS A STRING)
'T_site = Site period when required for considering interpolation for class D soils
'________________________________________________________________________________________________________________

'set entry to uppercase
    Site_subsoil_class = UCase(Site_subsoil_class)

    If Site_subsoil_class = "D" And T_site >= 0.6 And T_site <= 1.5 And T_1 >= 0.1 Then
        Loading_D_subsoil_interpolate_test = True
    Else
        Loading_D_subsoil_interpolate_test = False
    End If

End Function

Function Loading_C_T(T_1 As Double, Site_subsoil_class As String, Hazard_factor As Double, Return_period_factor As Double, _
                     Fault_distance As Variant, Optional ESM_case As Boolean = True, _
                     Optional D_subsoil_interpolate As Boolean = False, Optional T_site As Double = 1.5)
'Function to calculate C(T), the elastic site spectra

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_C_d_T(T_1,Site_subsoil_class,Hazard_factor,Return_period_factor,Fault_distance,ESM_case,D_subsoil_interpolate,T_site)
'T_1 = FIRST MODE PERIOD
'Site_subsoil_class = SITE SUBSOIL CLASS, i.e. A/B/C/D/E (ENTERED AS A STRING)
'Hazard_factor = HAZARD FACTOR Z
'Return_period_factor = RETURN PERIOD FACTOR Ru OR Rs
'Fault_distance = THE SHORTEST DISTANCE (IN kM's) FROM THE SITE TO THE NEAREST FAULT LISTED IN TABLE 3.6 OF NZS1170.5
'                 IF NOT RELEVANT USE "N/A" OR A NUMBER >= 20
'OPTIONAL - ESM_case is optional (ESM = Equivalent Static Method), when set to false the Response Spectrum Method (RSM)
'           shape factor will be calculated.
'OPTIONAL - D_subsoil_interpolation is optional - TRUE/FALSE - consider interpolation for class D soils
'OPTIONAL - T_site is optional - Site period when required for considering interpolation for class D soils
'           (note default value of 1.5 seconds equates to no interpolation)
'________________________________________________________________________________________________________________

    Dim C_h_T
    Dim N_T_D
    Dim Z_R_product

    'Spectral shape factor
    C_h_T = Loading_C_h_T(T_1, Site_subsoil_class, ESM_case, D_subsoil_interpolate, T_site)

    'Near fault factor
    N_T_D = Loading_N_T_D(T_1, Fault_distance, Return_period_factor)

    'check if Z x R > 0.7, if so limit product to 0.7 for calculating C_T
    If Hazard_factor * Return_period_factor > 0.7 Then
        Z_R_product = 0.7
    Else
        Z_R_product = Hazard_factor * Return_period_factor
    End If

    'Elastic site spectra
    Loading_C_T = C_h_T * Z_R_product * N_T_D

End Function

Function Loading_C_d_T(T_1 As Variant, Site_subsoil_class As String, Hazard_factor As Double, Return_period_factor As Double, _
                       Fault_distance As Variant, mu As Double, S_p As Double, Optional ESM_case As Boolean = True, _
                       Optional D_subsoil_interpolate As Boolean = False, Optional T_site As Double = 1.5) As Variant
'Function to calculate C_d(T), the seismic load coefficient for a single period T, or series of T periods arranged in a column range

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_C_d_T(T_1,Site_subsoil_class,Hazard_factor,Return_period_factor,Fault_distance,mu,S_p,ESM_case,D_subsoil_interpolate,T_site)
'T_1 = FIRST MODE PERIOD
'Site_subsoil_class = SITE SUBSOIL CLASS, i.e. A/B/C/D/E (ENTERED AS A STRING)
'Hazard_factor = HAZARD FACTOR Z
'Return_period_factor = RETURN PERIOD FACTOR Ru OR Rs
'Fault_distance = THE SHORTEST DISTANCE (IN kM's) FROM THE SITE TO THE NEAREST FAULT LISTED IN TABLE 3.6 OF NZS1170.5
'                 IF NOT RELEVANT USE "N/A" OR A NUMBER >= 20
'mu = DUCTILITY
'S_p = STRUCTURAL PERFORMANCE FACTOR
'OPTIONAL - ESM_case is optional (ESM = Equivalent Static Method), when set to false the Response Spectrum Method (RSM)
'           shape factor will be calculated.
'OPTIONAL - D_subsoil_interpolation is optional - TRUE/FALSE - consider interpolation for class D soils
'OPTIONAL - T_site is optional - Site period when required for considering interpolation for class D soils
'           (note default value of 1.5 seconds equates to no interpolation)
'________________________________________________________________________________________________________________

    Dim arr_temp As Double
    Dim k As Integer

    'convert to array
    'check if single value provided (i.e. one cell), and convert to 2D array
    If T_1.Rows.count = 1 Then
        T_1 = Array(T_1.Value2)
        arr_temp = T_1(0)
        ReDim T_1(1 To 1, 1 To 1)
        T_1(1, 1) = arr_temp
    Else
        'convert to 2D array
        T_1 = T_1.Value2
    End If

    Dim C_d_T
    ReDim C_d_T(LBound(T_1) To UBound(T_1), 1 To 1)

    For k = LBound(T_1) To UBound(T_1)
        C_d_T(k, 1) = Loading_C_d_T_intermediate(CDbl(T_1(k, 1)), Site_subsoil_class, Hazard_factor, Return_period_factor, _
                                                 Fault_distance, mu, S_p, ESM_case, D_subsoil_interpolate, T_site)
    Next k

    'return results
    Loading_C_d_T = C_d_T

End Function

Private Function Loading_C_d_T_intermediate(T_1 As Double, Site_subsoil_class As String, Hazard_factor As Double, Return_period_factor As Double, _
                                            Fault_distance As Variant, mu As Double, S_p As Double, Optional ESM_case As Boolean = True, _
                                            Optional D_subsoil_interpolate As Boolean = False, Optional T_site As Double = 1.5)
'Function to calculate C_d(T), the seismic load coefficient for a single period T

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_C_d_T_intermediate(T_1,Site_subsoil_class,Hazard_factor,Return_period_factor,Fault_distance,mu,S_p,ESM_case,D_subsoil_interpolate,T_site)
'T_1 = FIRST MODE PERIOD
'Site_subsoil_class = SITE SUBSOIL CLASS, i.e. A/B/C/D/E (ENTERED AS A STRING)
'Hazard_factor = HAZARD FACTOR Z
'Return_period_factor = RETURN PERIOD FACTOR Ru OR Rs
'Fault_distance = THE SHORTEST DISTANCE (IN kM's) FROM THE SITE TO THE NEAREST FAULT LISTED IN TABLE 3.6 OF NZS1170.5
'                 IF NOT RELEVANT USE "N/A" OR A NUMBER >= 20
'mu = DUCTILITY
'S_p = STRUCTURAL PERFORMANCE FACTOR
'OPTIONAL - ESM_case is optional (ESM = Equivalent Static Method), when set to false the Response Spectrum Method (RSM)
'           shape factor will be calculated.
'OPTIONAL - D_subsoil_interpolation is optional - TRUE/FALSE - consider interpolation for class D soils
'OPTIONAL - T_site is optional - Site period when required for considering interpolation for class D soils
'           (note default value of 1.5 seconds equates to no interpolation)
'________________________________________________________________________________________________________________

    Dim C_T
    Dim k_mu
    Dim C_d_T
    Dim Z_R_product

    'check if Z x R > 0.7, if so limit product to 0.7 for calculating C_T
    If Hazard_factor * Return_period_factor > 0.7 Then
        Z_R_product = 0.7
    Else
        Z_R_product = Hazard_factor * Return_period_factor
    End If

    'Elastic site spectra
    C_T = Loading_C_T(T_1, Site_subsoil_class, Hazard_factor, Return_period_factor, _
                      Fault_distance, ESM_case, D_subsoil_interpolate, T_site)

    'Inelastic spectrum scaling factor
    k_mu = Loading_k_mu(T_1, Site_subsoil_class, mu)

    'Horizontal design action coefficient
    C_d_T = C_T * S_p / k_mu

    'Check for minimum level design coefficient for ULS/SLS cases
    If C_d_T < Z_R_product / 20 + 0.02 * Return_period_factor Then
        C_d_T = Z_R_product / 20 + 0.02 * Return_period_factor
    End If
    If C_d_T < 0.03 * Return_period_factor Then
        C_d_T = 0.03 * Return_period_factor
    End If

    'return results
    Loading_C_d_T_intermediate = C_d_T

End Function

Function Loading_generate_period_range(period_range_limit As Double, period_step As Double, Optional T_site As Double = 1.5)
'Function to return array of period values at all points of interest for calculating values at all transition points
'and a distributed time step at 'period_step' intervals
'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_generate_period_range(period_range_limit,period_step)
'period_range_limit = upper limit of generated period sequence, starting at zero and ending at 'period_range_limit',
'                     and incoporating points of interest where variaous factors result in transition points in the
'                     resulting spectrum. The maximum step is taken as 'period_step' in the generated period sequence
'period_step = period increment in the generated period sequence
'   =Loading_generate_period_range(2,0.2) would return
'   0
'   0
'   0.1
'   0.1
'   0.2
'   0.3
'   0.4
'   0.4
'   0.56
'   0.6
'   0.8
'   1
'   1
'   1.2
'   1.4
'   1.5
'   1.6
'   1.8
'   2
'   2
'   2
'NOTE - there will be some duplicates as shown above, but this does not impact on plotting spectrum based on
'resulting period range
'________________________________________________________________________________________________________________

    Dim T_1_arr
    Dim num_steps As Long
    Dim num_points_of_interest As Long
    Dim i
    Dim steps

    num_steps = period_range_limit / period_step + 1
    num_points_of_interest = 12

    'resize results array for all regular intervals and points of interest
    ReDim steps(1 To num_steps + num_points_of_interest - 1)

    'create array of 'points of interest' where transitions or abrupt steps of various factors occurs
    T_1_arr = Array(0, 0.1 - 0.00000000001, 0.1, 0.3, 0.4, 0.56, 1, 1.5, 3, 4, 5, _
                    1 / (((3 / (1 + 0.5 * (T_site - 0.25))) / 2) ^ (1 / 0.75)) * 0.5)

    'populate regular period intervals starting at zero and max of 'period_range_limit'
    steps(1) = 0
    For i = 2 To num_steps - 1
        steps(i) = steps(i - 1) + period_step
    Next i
    steps(num_steps) = period_range_limit

    'populate points of interest into results array, limiting to max of 'period_range_limit'
    For i = 0 To num_points_of_interest - 1
        If T_1_arr(i) > period_range_limit Then
            steps(num_steps + i) = period_range_limit
        Else
            steps(num_steps + i) = T_1_arr(i)
        End If
    Next i

    'transpose array to single column of results
    steps = WorksheetFunction.Transpose(steps)
    'sort results into numerical order
    On Error GoTo skip_code
    'for excel 365
    steps = WorksheetFunction.Sort(steps)
    GoTo return_results
skip_code: 'for excel 2019
    steps = NZS1170_Seismic_Coefficient.array_quicksort_2D(steps)

return_results:
    'return results
    Loading_generate_period_range = steps

End Function

Function array_quicksort_2D(arr As Variant, Optional sort_column As Long = -1) As Variant
'Function that sorts a two-dimensional VBA array from smallest to largest using a divide and conquer algorithm
'sorting is undertaken based on the column specified, if no column is specified the lower bound column is used

    Dim lower_bound As Long
    Dim upper_bound As Long

    lower_bound = LBound(arr, 1)
    upper_bound = UBound(arr, 1)

    Call array_quicksort_2D_sub(arr, lower_bound, upper_bound, sort_column)

    'Return results
    array_quicksort_2D = arr

End Function

Private Sub array_quicksort_2D_sub(ByRef arr As Variant, lower_bound As Long, upper_bound As Long, Optional sort_column As Long = -1)
'Sub-procedure that sorts a two-dimensional VBA array from smallest to largest using a divide and conquer algorithm
'sorting is undertaken based on the column specified, if no column is specified the lower bound column is used

'called from array_quicksort_2D function, but can be used independantly of this incompassing function

    Dim temp_low As Long
    Dim temp_high As Long
    Dim pivot_value As Variant
    Dim temp_arr_row As Variant
    Dim temp_sort_column As Long

    If sort_column = -1 Then sort_column = LBound(arr, 2)
    temp_low = lower_bound
    temp_high = upper_bound
    pivot_value = arr((lower_bound + upper_bound) \ 2, sort_column)

    'Divide data in array
    While temp_low <= temp_high
        While arr(temp_low, sort_column) < pivot_value And temp_low < upper_bound
            temp_low = temp_low + 1
        Wend
        While pivot_value < arr(temp_high, sort_column) And temp_high > lower_bound
            temp_high = temp_high - 1
        Wend

        If temp_low <= temp_high Then    'swap rows if required
            ReDim temp_arr_row(LBound(arr, 2) To UBound(arr, 2))
            For temp_sort_column = LBound(arr, 2) To UBound(arr, 2)
                temp_arr_row(temp_sort_column) = arr(temp_low, temp_sort_column)
                arr(temp_low, temp_sort_column) = arr(temp_high, temp_sort_column)
                arr(temp_high, temp_sort_column) = temp_arr_row(temp_sort_column)
            Next temp_sort_column
            Erase temp_arr_row
            temp_low = temp_low + 1
            temp_high = temp_high - 1
        End If
    Wend

    'Sort data in array in iterative process
    If (lower_bound < temp_high) Then Call array_quicksort_2D_sub(arr, lower_bound, temp_high, sort_column)
    If (temp_low < upper_bound) Then Call array_quicksort_2D_sub(arr, temp_low, upper_bound, sort_column)

End Sub

Function Loading_C_d_T_part(T_part As Variant, Site_subsoil_class As String, Hazard_factor As Double, Return_period_factor As Double, _
                            R_p As Double, mu_part As Double, h_i As Double, h_n As Double, Optional C_ph_code_values As Boolean = True) As Variant
'Function to calculate the parts and components seismic coefficient

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_C_d_T_intermediate_part(T_part,Site_subsoil_class,Hazard_factor,Return_period_factor,R_p,mu_part,h_i,h_n,C_ph_code_values)
'T_part = PART FIRST MODE PERIOD
'Site_subsoil_class = SITE SUBSOIL CLASS, i.e. A/B/C/D/E (ENTERED AS A STRING)
'Hazard_factor = HAZARD FACTOR Z
'Return_period_factor = RETURN PERIOD FACTOR Ru OR Rs
'mu_part = DUCTILITY OF PART
'h_i = HEIGHT OF ATTACHMENT OF PART
'h_n = HEIGHT FROM THE BASE OF THE STRUCTURE TO THE UPPERMOST SEISMIC WEIGHT OR MASS
'OPTIONAL - C_ph_code_values is optional, when set to TRUE the value is interpolated from the NZS1170.5 TABLE C8.3
'           rounded values. When set to FALSE, the values are calculated based on first principles approach
'________________________________________________________________________________________________________________

    Dim arr_temp As Double
    Dim k As Integer

    'convert to array
    'check if single value provided (i.e. one cell), and convert to 2D array
    If T_part.Rows.count = 1 Then
        T_part = Array(T_part.Value2)
        arr_temp = T_part(0)
        ReDim T_part(1 To 1, 1 To 1)
        T_part(1, 1) = arr_temp
    Else
        'convert to 2D array
        T_part = T_part.Value2
    End If

    Dim C_d_T_part
    ReDim C_d_T_part(LBound(T_part) To UBound(T_part), 1 To 1)

    For k = LBound(T_part) To UBound(T_part)
        C_d_T_part(k, 1) = Loading_C_d_T_intermediate_part(T_part(k, 1), Site_subsoil_class, Hazard_factor, Return_period_factor, _
                                                           R_p, mu_part, h_i, h_n, C_ph_code_values)
    Next k

    'return results
    Loading_C_d_T_part = C_d_T_part

End Function

Private Function Loading_C_d_T_intermediate_part(T_part As Variant, Site_subsoil_class As String, Hazard_factor As Double, Return_period_factor As Double, _
                                                 R_p As Double, mu_part As Double, h_i As Double, h_n As Double, Optional C_ph_code_values As Boolean = True)
'Function to calculate C_d(T) for parts and components loading, the parts seismic load coefficient for a single period T

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_C_d_T_intermediate_part(T_part,Site_subsoil_class,Hazard_factor,Return_period_factor,R_p,mu_part,h_i,h_n,C_ph_code_values)
'T_part = PART FIRST MODE PERIOD
'Site_subsoil_class = SITE SUBSOIL CLASS, i.e. A/B/C/D/E (ENTERED AS A STRING)
'Hazard_factor = HAZARD FACTOR Z
'Return_period_factor = RETURN PERIOD FACTOR Ru OR Rs
'mu_part = DUCTILITY OF PART
'h_i = HEIGHT OF ATTACHMENT OF PART
'h_n = HEIGHT FROM THE BASE OF THE STRUCTURE TO THE UPPERMOST SEISMIC WEIGHT OR MASS
'OPTIONAL - C_ph_code_values is optional, when set to TRUE the value is interpolated from the NZS1170.5 TABLE C8.3
'           rounded values. When set to FALSE, the values are calculated based on first principles approach
'________________________________________________________________________________________________________________

    Dim C_T0_part
    Dim C_Hi_part
    Dim C_i_T_part
    Dim C_ph_part
    Dim C_d_T_part
    Dim Z_R_product

    'check if Z x R > 0.7, if so limit product to 0.7 for calculating C_T
    If Hazard_factor * Return_period_factor > 0.7 Then
        Z_R_product = 0.7
    Else
        Z_R_product = Hazard_factor * Return_period_factor
    End If

    'Elastic site spectra for T=0 seconds for modal response spectrum method
    C_T0_part = Loading_C_T(0, Site_subsoil_class, Hazard_factor, Return_period_factor, "N/A", False)

    'Determine floor height coefficient
    C_Hi_part = Loading_C_Hi(h_i, h_n)

    'determine part or component spectral shape coefficient
    C_i_T_part = Loading_C_i_T_p(T_part)

    'determine part horizontal response factor from first principles calculation, this will result in slightly different values as
    'the results are not rounded like the code values (refer commentary for tabulated values)
    C_ph_part = Loading_C_ph_part(mu_part, C_ph_code_values)

    'Horizontal design action coefficient
    C_d_T_part = C_T0_part * C_Hi_part * C_i_T_part * C_ph_part * R_p

    'Check for maximum level design coefficient of 3.6
    If C_d_T_part > 3.6 Then C_d_T_part = 3.6

    'return results
    Loading_C_d_T_intermediate_part = C_d_T_part

End Function

Function Loading_C_ph_part(mu_part As Double, Optional C_ph_code_values As Boolean = True)
'Function to return the part response factor, calculated from first priciples based on the TABLE C8.3 in the commentary of NZS1170.5
'or from TABLE C8.3 values

'refer following Loading_k_mu_part and Loading_Sp_1170 for basis of the first principles calculation

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_C_ph_part(mu_part,C_ph_code_values)
'mu_part = DUCTILITY OF PART
'OPTIONAL - C_ph_code_values is optional, when set to TRUE the value is interpolated from the NZS1170.5 TABLE C8.3
'           rounded values. When set to FALSE, the values are calculated based on first principles approach
'________________________________________________________________________________________________________________

    Dim mu_part_array
    Dim C_p_round_array
    Dim lower_bound_index As Long
    Dim upper_bound_index As Long
    Dim i As Long

    If C_ph_code_values Then
        'determine part horizontal response factor from code rounded values in TABLE C8.3
        mu_part_array = Array(1, 1.25, 1.5, 1.75, 2, 3, 6)
        C_p_round_array = Array(1, 0.85, 0.75, 0.65, 0.55, 0.45, 0.45)
        'interpolate for C_ph value
        'determine upper bound index
        For i = LBound(mu_part_array) To UBound(mu_part_array)
            'determine in what band the actual ductility lies
            If mu_part <= mu_part_array(i) Then
                upper_bound_index = i
                GoTo skip_code1
            End If
        Next i
skip_code1:
        'determine lower bound index
        For i = UBound(mu_part_array) To LBound(mu_part_array) Step -1
            'determine in what band the actual ductility lies
            If mu_part >= mu_part_array(i) Then
                lower_bound_index = i
                GoTo skip_code2
            End If
        Next i
skip_code2:

        'interpolate for code C_ph value based on mu_part
        If lower_bound_index = upper_bound_index Then
            Loading_C_ph_part = C_p_round_array(lower_bound_index)
        Else
            Loading_C_ph_part = C_p_round_array(lower_bound_index) * _
                                (1 - (mu_part - mu_part_array(lower_bound_index)) / (mu_part_array(upper_bound_index) - mu_part_array(lower_bound_index))) + _
                                C_p_round_array(upper_bound_index) * _
                                ((mu_part - mu_part_array(lower_bound_index)) / (mu_part_array(upper_bound_index) - mu_part_array(lower_bound_index)))
        End If

    Else
        'determin part horizontal response factor from first principles
        Loading_C_ph_part = Loading_S_p_1170(mu_part) / Loading_k_mu_part(mu_part)
    End If

End Function

Function Loading_C_Hi(h_i As Double, h_n As Double)
'Function to calculate C_Hi the floor height coefficient for parts & components loading

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_C_Hi(h_i,h_n)
'h_i = HEIGHT OF ATTACHMENT OF PART
'h_n = HEIGHT FROM THE BASE OF THE STRUCTURE TO THE UPPERMOST SEISMIC WEIGHT OR MASS
'________________________________________________________________________________________________________________

    Dim C_Hi1 As Double
    Dim C_Hi2 As Double
    Dim C_Hi3 As Double

    C_Hi1 = 1 + h_i / 6
    C_Hi2 = 1 + 10 * (h_i / h_n)
    C_Hi3 = 3

    If h_i < 12 Or h_i < 0.2 * h_n Then
        Loading_C_Hi = WorksheetFunction.Min(C_Hi1, C_Hi2)
    ElseIf h_i < 0.2 * h_n Then
        Loading_C_Hi = C_Hi2
    Else
        Loading_C_Hi = C_Hi3
    End If

End Function

Function Loading_C_i_T_p(T_part As Variant)
'Function to calculate Ci(Tp) the part spectral shape coefficient for parts & components loading

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_C_i_T_p(T_part)
'T_part = PART FIRST MODE PERIOD
'________________________________________________________________________________________________________________

    If T_part <= 0.75 Then
        Loading_C_i_T_p = 2
    ElseIf T_part >= 1.5 Then
        Loading_C_i_T_p = 0.5
    Else
        Loading_C_i_T_p = 2 * (1.75 - T_part)
    End If

End Function

Function Loading_k_mu_part(mu_part As Double)
'Function to return the k_mu factor appropriate for parts, determined from values in TABLE C8.3 in NZS1170.5 commentary.

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_k_mu_part(mu_part)
'mu_part = DUCTILITY OF PART

'NOTE - the calculated values in TABLE C8.3 are consistent with the calculation at T = 0 seconds,
'       using any soil classification except "E" and the following relationship
'       k_mu_part = (k_mu - 1) / 2 + 1
'       This will result in the values in TABLE C8.3 being actually calculated without any rounding
'________________________________________________________________________________________________________________

    Dim Site_subsoil_class As String

    Site_subsoil_class = "A"    'could use any value here apart from "E"

    Loading_k_mu_part = (Loading_k_mu(0, Site_subsoil_class, mu_part) - 1) / 2 + 1

End Function

Function Loading_S_p_1170(mu As Double)
'Function to return an Sp factor in accordance with NZS1170.5 CL4.4.2

'________________________________________________________________________________________________________________
'USAGE
'________________________________________________________________________________________________________________
'=Loading_S_p_1170(mu_part)
'mu = DUCTILITY
'________________________________________________________________________________________________________________

    Loading_S_p_1170 = 1.3 - 0.3 * mu

    'check lower limit if S_p < 0.7
    If Loading_S_p_1170 < 0.7 Then Loading_S_p_1170 = 0.7

End Function
