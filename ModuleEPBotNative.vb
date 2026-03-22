Imports System.Runtime.InteropServices
Imports System.Text

'''
''' Wrapper P/Invoke dla natywnej biblioteki AOT EPBot.dll (Windows x64).
''' Sygnatury zgodne z epbot.h.
'''
''' Konwencje z nagłówka:
'''   - Stringi: UTF-8, null-terminated
'''   - Wyjściowe stringi: caller-provided buffer + buffer_size
'''   - Tablice stringów: newline-separated w jednym buforze
'''   - Boolean: int (0=false, nonzero=true)
'''   - handle: nieprzezroczysty wskaźnik z epbot_create()
'''
Module ModuleEPBotNative

    Private Const DLL As String = "EPBot.dll"
    Private Const BUF As Integer = 4096   ' rozmiar domyślnego bufora

    ' ── Błędy ────────────────────────────────────────────────────
    Public Const EPBOT_OK As Integer = 0
    Public Const EPBOT_ERR_NULL_HANDLE As Integer = -1
    Public Const EPBOT_ERR_EXCEPTION As Integer = -2
    Public Const EPBOT_ERR_BUFFER_SMALL As Integer = -3

    ' ════════════════════════════════════════════════════════════
    '  Surowe deklaracje P/Invoke
    ' ════════════════════════════════════════════════════════════

    ' ── Lifecycle ────────────────────────────────────────────────

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_create() As IntPtr
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Sub epbot_destroy(instance As IntPtr)
    End Sub

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_last_error() As IntPtr   ' const char* — nie zwalniać
    End Function

    ' ── Core bidding ─────────────────────────────────────────────

    ' longer: karty jako "AKQJ\nT98\n765\n432" (C.D.H.S, newline-separated)
    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_new_hand(instance As IntPtr,
                                    player_position As Integer,
                                    <MarshalAs(UnmanagedType.LPUTF8Str)> longer As String,
                                    dealer As Integer,
                                    vulnerability As Integer,
                                    repeating As Integer,
                                    b_playing As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_bid(instance As IntPtr) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_bid(instance As IntPtr,
                                   spare As Integer,
                                   new_value As Integer,
                                   <MarshalAs(UnmanagedType.LPUTF8Str)> str_alert As String) As Integer
    End Function

    ' bids: odzywki newline-separated
    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_arr_bids(instance As IntPtr,
                                        <MarshalAs(UnmanagedType.LPUTF8Str)> bids As String) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_interpret_bid(instance As IntPtr, bid_code As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_ask(instance As IntPtr) As Integer
    End Function

    ' ── Conventions ──────────────────────────────────────────────

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_conventions(instance As IntPtr,
                                           site As Integer,
                                           <MarshalAs(UnmanagedType.LPUTF8Str)> convention As String) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_conventions(instance As IntPtr,
                                           site As Integer,
                                           <MarshalAs(UnmanagedType.LPUTF8Str)> convention As String,
                                           value As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_system_type(instance As IntPtr, system_number As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_system_type(instance As IntPtr, system_number As Integer, value As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_opponent_type(instance As IntPtr, system_number As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_opponent_type(instance As IntPtr, system_number As Integer, value As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_convention_index(instance As IntPtr,
                                            <MarshalAs(UnmanagedType.LPUTF8Str)> name As String) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_convention_name(instance As IntPtr, index As Integer,
                                           buffer As Byte(), buffer_size As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_convention_name(instance As IntPtr, index As Integer,
                                               buffer As Byte(), buffer_size As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_selected_conventions(instance As IntPtr,
                                                buffer As Byte(), buffer_size As Integer,
                                                ByRef count_out As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_system_name(instance As IntPtr, system_number As Integer,
                                       buffer As Byte(), buffer_size As Integer) As Integer
    End Function

    ' ── Scoring & settings ───────────────────────────────────────

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_scoring(instance As IntPtr) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_scoring(instance As IntPtr, value As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_playing_skills(instance As IntPtr) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_playing_skills(instance As IntPtr, value As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_defensive_skills(instance As IntPtr) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_defensive_skills(instance As IntPtr, value As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_licence(instance As IntPtr) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_licence(instance As IntPtr, value As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_bcalconsole_path(instance As IntPtr,
                                                buffer As Byte(), buffer_size As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_bcalconsole_path(instance As IntPtr,
                                                <MarshalAs(UnmanagedType.LPUTF8Str)> path As String) As Integer
    End Function

    ' ── State queries ────────────────────────────────────────────

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_position(instance As IntPtr) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_dealer(instance As IntPtr) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_vulnerability(instance As IntPtr) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_version(instance As IntPtr) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_copyright(instance As IntPtr,
                                     buffer As Byte(), buffer_size As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_last_epbot_error(instance As IntPtr,
                                                buffer As Byte(), buffer_size As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_str_bidding(instance As IntPtr,
                                           buffer As Byte(), buffer_size As Integer) As Integer
    End Function

    ' ── Info / meaning ───────────────────────────────────────────

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_info_alerting(instance As IntPtr, k As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_info_alerting(instance As IntPtr, k As Integer, value As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_info_meaning(instance As IntPtr, k As Integer,
                                            buffer As Byte(), buffer_size As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_info_meaning(instance As IntPtr, k As Integer,
                                            <MarshalAs(UnmanagedType.LPUTF8Str)> value As String) As Integer
    End Function

    ' ── Card play ────────────────────────────────────────────────

    ' arr_cards: karty dziadka newline-separated
    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_dummy(instance As IntPtr,
                                     dummy As Integer,
                                     <MarshalAs(UnmanagedType.LPUTF8Str)> arr_cards As String,
                                     all_data As Integer,
                                     ByRef without_final_length As Byte) As Integer
    End Function

    ' current_longers: 0 lub 1; wynik: newline-separated suits
    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_arr_suits(instance As IntPtr,
                                         current_longers As Integer,
                                         buffer As Byte(), buffer_size As Integer,
                                         ByRef count_out As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_hand(instance As IntPtr,
                                    player_position As Integer,
                                    buffer As Byte(), buffer_size As Integer,
                                    ByRef count_out As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_lead(instance As IntPtr,
                                    force_lead As Integer,
                                    buffer As Byte(), buffer_size As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_set_lead(instance As IntPtr,
                                    <MarshalAs(UnmanagedType.LPUTF8Str)> played_card As String) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_cards(instance As IntPtr,
                                     buffer As Byte(), buffer_size As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_probable_level(instance As IntPtr, strain As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_info_feature(instance As IntPtr, position As Integer,
                                            buffer As Byte(), buffer_size As Integer,
                                            ByRef count_out As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_info_min_length(instance As IntPtr, position As Integer,
                                               buffer As Byte(), buffer_size As Integer,
                                               ByRef count_out As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_info_max_length(instance As IntPtr, position As Integer,
                                               buffer As Byte(), buffer_size As Integer,
                                               ByRef count_out As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_info_probable_length(instance As IntPtr, position As Integer,
                                                    buffer As Byte(), buffer_size As Integer,
                                                    ByRef count_out As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_info_honors(instance As IntPtr, position As Integer,
                                           buffer As Byte(), buffer_size As Integer,
                                           ByRef count_out As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_info_suit_power(instance As IntPtr, position As Integer,
                                               buffer As Byte(), buffer_size As Integer,
                                               ByRef count_out As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_info_strength(instance As IntPtr, position As Integer,
                                             buffer As Byte(), buffer_size As Integer,
                                             ByRef count_out As Integer) As Integer
    End Function

    <DllImport(DLL, CallingConvention:=CallingConvention.Cdecl)>
    Private Function epbot_get_info_stoppers(instance As IntPtr, position As Integer,
                                             buffer As Byte(), buffer_size As Integer,
                                             ByRef count_out As Integer) As Integer
    End Function

    ' ════════════════════════════════════════════════════════════
    '  Pomocnicze
    ' ════════════════════════════════════════════════════════════

    ''' <summary>Odczytuje bufor bajtów UTF-8 jako String (do pierwszego \0).</summary>
    Private Function BufToString(buf As Byte()) As String
        Dim nul As Integer = Array.IndexOf(buf, CByte(0))
        If nul < 0 Then nul = buf.Length
        Return Encoding.UTF8.GetString(buf, 0, nul)
    End Function

    ''' <summary>Konwertuje tablicę String() na jeden string newline-separated.</summary>
    Private Function SuitsToLonger(suits As String()) As String
        ' suits(0)=C suits(1)=D suits(2)=H suits(3)=S
        Return String.Join(vbLf, suits)
    End Function

    ''' <summary>Pobiera ostatni błąd FFI jako String.</summary>
    Public Function GetLastError() As String
        Dim ptr As IntPtr = epbot_get_last_error()
        If ptr = IntPtr.Zero Then Return String.Empty
        Return Marshal.PtrToStringUTF8(ptr)
    End Function

    ' ════════════════════════════════════════════════════════════
    '  Klasa opakowująca — identyczny interfejs jak EPBot8739.EPBot
    ' ════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Wrapper natywnej biblioteki EPBot (AOT DLL) z interfejsem identycznym
    ''' jak EPBot8739.EPBot — pozostały kod projektu nie wymaga zmian.
    ''' </summary>
    Public Class EPBot
        Implements IDisposable

        Private ReadOnly _h As IntPtr
        Private _disposed As Boolean = False

        Public Sub New()
            _h = epbot_create()
            If _h = IntPtr.Zero Then
                Throw New InvalidOperationException(
                    "epbot_create() zwróciło null. Sprawdź czy EPBot.dll jest w katalogu exe.")
            End If
        End Sub

        ' ── Zarządzanie ręką ──────────────────────────────────────

        ''' <summary>
        ''' Inicjalizuje rękę gracza.
        ''' suits(): indeks 0=C, 1=D, 2=H, 3=S — zgodnie z C_CLUBS..C_SPADES.
        ''' </summary>
        Public Sub new_hand(position As Integer, suits As String(),
                            dealer As Integer, vulnerable As Integer)
            Dim longer As String = SuitsToLonger(suits)
            Dim rc As Integer = epbot_new_hand(_h, position, longer, dealer, vulnerable, 0, 0)
            CheckRc(rc, "epbot_new_hand")
        End Sub

        ' ── Licytacja ─────────────────────────────────────────────

        Public Function get_bid() As Integer
            Dim rc As Integer = epbot_get_bid(_h)
            If rc < 0 Then Throw New InvalidOperationException($"epbot_get_bid error {rc}: {GetLastError()}")
            Return rc
        End Function

        ''' <summary>Rozgłasza odzywkę (str_alert opcjonalny).</summary>
        Public Sub set_bid(position As Integer, bid As Integer,
                           Optional alert As String = "")
            Dim rc As Integer = epbot_set_bid(_h, position, bid, If(alert, ""))
            CheckRc(rc, "epbot_set_bid")
        End Sub

        ''' <summary>Ustawia pełną historię licytacji z tablicy arr_bids.</summary>
        Public Sub set_arr_bids(bids As String())
            Dim filled As New System.Collections.Generic.List(Of String)
            For Each b As String In bids
                If String.IsNullOrEmpty(b) Then Exit For
                filled.Add(Left(b, 2))  ' tylko kod numeryczny, bez alertu
            Next
            Dim rc As Integer = epbot_set_arr_bids(_h, String.Join(vbLf, filled))
            CheckRc(rc, "epbot_set_arr_bids")
        End Sub

        ' ── Rozgrywka ─────────────────────────────────────────────

        ''' <summary>Ustawia dziadka i jego karty.</summary>
        Public Sub set_dummy(dummy As Integer, suits As String())
            Dim longer As String = SuitsToLonger(suits)
            Dim wfl As Byte = 0
            Dim rc As Integer = epbot_set_dummy(_h, dummy, longer, 0, wfl)
            CheckRc(rc, "epbot_set_dummy")
        End Sub

        ''' <summary>
        ''' Zwraca tablicę stringów (karty wszystkich graczy) jako newline-separated,
        ''' zgodnie z układem EPBot8739: 4 graczy × 4 kolory.
        ''' </summary>
        Public Function get_arr_suits() As String()
            Dim outBuf(BUF - 1) As Byte
            Dim count As Integer = 0
            Dim rc As Integer = epbot_get_arr_suits(_h, 0, outBuf, BUF, count)
            CheckRc(rc, "epbot_get_arr_suits")
            Return BufToString(outBuf).Split(vbLf)
        End Function

        ' ── Informacje o odzywce ──────────────────────────────────

        Public Function info_alerting(position As Integer) As Boolean
            Return epbot_get_info_alerting(_h, position) <> 0
        End Function

        Public Function info_meaning(position As Integer) As String
            Dim outBuf(BUF - 1) As Byte
            Dim rc As Integer = epbot_get_info_meaning(_h, position, outBuf, BUF)
            If rc < 0 Then Return String.Empty
            Return BufToString(outBuf)
        End Function

        Public Function get_str_bidding() As String
            Dim outBuf(BUF - 1) As Byte
            Dim rc As Integer = epbot_get_str_bidding(_h, outBuf, BUF)
            If rc < 0 Then Return String.Empty
            Return BufToString(outBuf)
        End Function

        Public Function interpret_bid(bid_code As Integer) As Integer
            Return epbot_interpret_bid(_h, bid_code)
        End Function

        Public Function ask() As Integer
            Return epbot_ask(_h)
        End Function

        Public ReadOnly Property LastError As String
            Get
                Return get_last_epbot_error()
            End Get
        End Property

        ' ── State queries ─────────────────────────────────────────

        Public Function get_position() As Integer
            Return epbot_get_position(_h)
        End Function

        Public Function get_dealer() As Integer
            Return epbot_get_dealer(_h)
        End Function

        Public Function get_vulnerability() As Integer
            Return epbot_get_vulnerability(_h)
        End Function

        Public Function get_last_epbot_error() As String
            Dim outBuf(511) As Byte
            Dim rc As Integer = epbot_get_last_epbot_error(_h, outBuf, 512)
            If rc < 0 Then Return String.Empty
            Return BufToString(outBuf)
        End Function

        ' ── Analiza ───────────────────────────────────────────────

        Public Function get_probable_level(strain As Integer) As Integer
            Return epbot_get_probable_level(_h, strain)
        End Function

        ' ── Card play ─────────────────────────────────────────────

        Public Function get_lead(force_lead As Boolean) As String
            Dim outBuf(255) As Byte
            Dim rc As Integer = epbot_get_lead(_h, If(force_lead, 1, 0), outBuf, 256)
            If rc < 0 Then Return String.Empty
            Return BufToString(outBuf)
        End Function

        Public Sub set_lead(played_card As String)
            CheckRc(epbot_set_lead(_h, played_card), "epbot_set_lead")
        End Sub

        Public Function get_cards() As String
            Dim outBuf(BUF - 1) As Byte
            Dim rc As Integer = epbot_get_cards(_h, outBuf, BUF)
            If rc < 0 Then Return String.Empty
            Return BufToString(outBuf)
        End Function

        Public Function get_hand(player_position As Integer) As String()
            Dim outBuf(BUF - 1) As Byte
            Dim count As Integer = 0
            Dim rc As Integer = epbot_get_hand(_h, player_position, outBuf, BUF, count)
            If rc < 0 Then Return New String() {}
            Return BufToString(outBuf).Split(vbLf)
        End Function

        ' ── Info arrays ───────────────────────────────────────────

        Public Function info_feature(position As Integer) As Integer()
            Return GetInfoIntArray(position, AddressOf epbot_get_info_feature)
        End Function

        Public Function info_min_length(position As Integer) As Integer()
            Return GetInfoIntArray(position, AddressOf epbot_get_info_min_length)
        End Function

        Public Function info_max_length(position As Integer) As Integer()
            Return GetInfoIntArray(position, AddressOf epbot_get_info_max_length)
        End Function

        Public Function info_probable_length(position As Integer) As Integer()
            Return GetInfoIntArray(position, AddressOf epbot_get_info_probable_length)
        End Function

        Public Function info_honors(position As Integer) As Integer()
            Return GetInfoIntArray(position, AddressOf epbot_get_info_honors)
        End Function

        Public Function info_suit_power(position As Integer) As Integer()
            Return GetInfoIntArray(position, AddressOf epbot_get_info_suit_power)
        End Function

        Public Function info_strength(position As Integer) As Integer()
            Return GetInfoIntArray(position, AddressOf epbot_get_info_strength)
        End Function

        Public Function info_stoppers(position As Integer) As Integer()
            Return GetInfoIntArray(position, AddressOf epbot_get_info_stoppers)
        End Function

        ' helper dla int-array getterów
        Private Delegate Function InfoIntArrayFunc(instance As IntPtr, position As Integer,
                                                   buffer As Byte(), buffer_size As Integer,
                                                   ByRef count_out As Integer) As Integer

        Private Function GetInfoIntArray(position As Integer, fn As InfoIntArrayFunc) As Integer()
            Dim rawBuf(BUF - 1) As Byte
            Dim count As Integer = 0
            Dim rc As Integer = fn(_h, position, rawBuf, BUF, count)
            If rc < 0 OrElse count = 0 Then Return New Integer() {}
            Dim result(count - 1) As Integer
            System.Buffer.BlockCopy(rawBuf, 0, result, 0, count * 4)
            Return result
        End Function

        ' ── Właściwości systemu ───────────────────────────────────

        ''' <summary>Typ systemu licytacyjnego (0=NS, 1=WE): T_21GF, T_SAYC itd.</summary>
        Public Property system_type(team As Integer) As Integer
            Get
                Return epbot_get_system_type(_h, team)
            End Get
            Set(value As Integer)
                CheckRc(epbot_set_system_type(_h, team, value), "epbot_set_system_type")
            End Set
        End Property

        ''' <summary>Sposób punktacji: 0=MP, 1=IMP.</summary>
        Public Property scoring As Integer
            Get
                Return epbot_get_scoring(_h)
            End Get
            Set(value As Integer)
                CheckRc(epbot_set_scoring(_h, value), "epbot_set_scoring")
            End Set
        End Property

        ''' <summary>Konwencja dla danego zespołu i nazwy.</summary>
        Public Property conventions(team As Integer, name As String) As Integer
            Get
                Return epbot_get_conventions(_h, team, name)
            End Get
            Set(value As Integer)
                CheckRc(epbot_set_conventions(_h, team, name, value), "epbot_set_conventions")
            End Set
        End Property

        ' ── Info ─────────────────────────────────────────────────

        Public ReadOnly Property Version As Integer
            Get
                Return epbot_version(_h)
            End Get
        End Property

        Public ReadOnly Property Copyright As String
            Get
                Dim outBuf(511) As Byte
                epbot_copyright(_h, outBuf, 512)
                Return BufToString(outBuf)
            End Get
        End Property

        ' ── Obsługa błędów ────────────────────────────────────────

        Private Sub CheckRc(rc As Integer, name As String)
            If rc < 0 Then
                Dim detail As String = String.Empty
                Dim ebuf(511) As Byte
                If epbot_get_last_epbot_error(_h, ebuf, 512) >= 0 Then
                    detail = BufToString(ebuf)
                End If
                If String.IsNullOrEmpty(detail) Then detail = GetLastError()
                Throw New InvalidOperationException($"{name} zwróciło {rc}: {detail}")
            End If
        End Sub

        ' ── IDisposable ───────────────────────────────────────────

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                epbot_destroy(_h)
                _disposed = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overrides Sub Finalize()
            Dispose(False)
            MyBase.Finalize()
        End Sub

    End Class

End Module
