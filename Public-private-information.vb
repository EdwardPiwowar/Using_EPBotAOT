Imports System
Imports System.Runtime.CompilerServices

Module Program
    Sub Main(args As String())
        Dim i As Long, j As Integer, k As Integer, n As Integer, asker As Integer, trump As Integer
        Dim str_hand As String, str_error As String
        Dim info_min() As Integer
        Dim info_max() As Integer
        'Dim honors() As Integer
        ReDim hand(3)
        ReDim arr_bids(127)
        ReDim arr_leads(63)
        ReDim arr_leaders(51)
        For k = 0 To 3
            ReDim hand(k).suit(3)
        Next
        For k = 0 To 3
            Player(k) = New ModuleEPBotNative.EPBot
        Next k
        Console.WriteLine("Public private information")
        set_board()
        set_dealers()
        set_vulnerability()
        set_strain_mark()

        vulnerable = 3
        deal = 1
        dealer = 3

        ReDim arr_bids(63)
        arr_bids(0) = "00"
        arr_bids(1) = "10"
        arr_bids(2) = "00"
        arr_bids(3) = "12"
        arr_bids(4) = "00"
        arr_bids(5) = "14"
        arr_bids(6) = "00"
        arr_bids(7) = "16"
        arr_bids(8) = "00"
        arr_bids(9) = "18"
        arr_bids(10) = "00"
        arr_bids(11) = "19"
        arr_bids(12) = "00"
        arr_bids(13) = "21"
        arr_bids(14) = "00"
        arr_bids(15) = "24"
        arr_bids(16) = "00"
        arr_bids(17) = "28"
        'arr_bids(18) = "00"

        k = 0
        With Player(k)
            asker = -1
            trump = C_NT
            'IMPORTANT - it is required to establish a system for both lines
            .system_type(0) = T_21GF
            .system_type(1) = T_21GF
            .conventions(0, "Blackwood 1430") = True
            .conventions(1, "Blackwood 1430") = True
            .conventions(0, "Blackwood without K And Q") = False
            .conventions(1, "Blackwood without K And Q") = False

            'set hand - it is essential for a player to know his cards, position and vulnerability
            hand(k).suit(3) = "AQT8"
            hand(k).suit(2) = "AK"
            hand(k).suit(1) = "AK6"
            hand(k).suit(0) = "A986"
            Console.WriteLine("Dll version " & .version)
            .new_hand(k, hand(k).suit, dealer, vulnerable)
            str_error = .LastError
            If str_error <> "" Then
                Console.WriteLine("Error " & str_error)
                Exit Sub
            End If
            Console.WriteLine("Position=" & .get_Position)
            Console.WriteLine("Dealer=" & .get_Dealer)
            Console.WriteLine("Vulnerability=" & .get_Vulnerability)
            Console.WriteLine(.get_Cards)
            ''set the entire auction
            .set_arr_bids(arr_bids)
            Console.WriteLine(.get_str_bidding)
            Console.WriteLine("")
            For n = 2 To 7 Step 4
                Console.WriteLine("Player " & k & " asks about position " & n)
                info = .info_feature(n)
                Console.WriteLine("HCP " & info(402) & "-" & info(403))
                info_min = .info_min_length(n)
                info_max = .info_max_length(n)
                Console.WriteLine("Length")
                For i = 3 To 0 Step -1
                    Console.Write(strain_mark(i) & " " & info_min(i) & "-" & info_max(i) & vbTab)
                Next i
                Console.WriteLine("")
                If info(425) > 0 Then
                    asker = n
                    Console.WriteLine("asker=" & asker)
                    trump = info(424)
                    Console.WriteLine("trump=" & trump)
                End If
                info = .info_stoppers(n)
                Console.WriteLine("Stoppers")
                For i = 3 To 0 Step -1
                    Console.Write(strain_mark(i) & " " & info(i) & vbTab)
                Next i
                Console.WriteLine("")
                Console.WriteLine("")
                info = .info_feature(n)
                Console.WriteLine("A=" & info(406))
                Console.WriteLine("K=" & info(407))
                Console.WriteLine("Q=" & info(319)) '-1 - no trump Q, 0 not set, 1 - trump Q 
                Console.WriteLine("")
            Next n
            arr_suits = .get_arr_suits
            print_hand_diagram()

        End With
        Console.WriteLine("")
        Console.ReadKey()
    End Sub
End Module
