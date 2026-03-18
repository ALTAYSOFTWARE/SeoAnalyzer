Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Net
Imports System.Net.Http
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Shapes
Imports HtmlAgilityPack
Imports Microsoft.VisualBasic

Partial Public Class MainWindow

    ' ============================================================
    '  MODEL SINIFLARI
    ' ============================================================
    Public Class ChecklistItem
        Public Property Title As String
        Public Property Detail As String
        Public Property Icon As String
        Public Property Badge As String
        Public Property StatusColor As SolidColorBrush
        Public Property BadgeColor As SolidColorBrush
        Public Property BadgeText As SolidColorBrush
    End Class

    Public Class HeadingItem
        Public Property Tag As String
        Public Property Text As String
        Public Property TagColor As SolidColorBrush
    End Class

    Public Class LinkItem
        Public Property Url As String
        Public Property RawUrl As String
        Public Property AnchorText As String
        Public Property LinkType As String
        Public Property TypeColor As SolidColorBrush
        Public Property StatusCode As String
        Public Property StatusColor As SolidColorBrush
    End Class

    Public Class ImageItem
        Public Property Src As String
        Public Property AltText As String
        Public Property AltStatus As String
        Public Property AltColor As SolidColorBrush
        Public Property AltBadgeColor As SolidColorBrush
    End Class

    Public Class KeywordItem
        Public Property Word As String
        Public Property CountText As String
        Public Property Percent As Double
    End Class

    Public Class CompareItem
        Public Property Label As String
        Public Property Val1 As String
        Public Property Val2 As String
        Public Property Winner As String
        Public Property WinColor As SolidColorBrush
        Public Property WinFg As SolidColorBrush
    End Class

    Public Class HistoryItem
        Public Property Url As String
        Public Property AnalysisDate As String
        Public Property Score As String
        Public Property Speed As String
        Public Property ScoreColor As SolidColorBrush
        Public Property ScoreFg As SolidColorBrush
    End Class

    Public Class TechItem
        Public Property Label As String
        Public Property Value As String
        Public Property Status As String
        Public Property BadgeColor As SolidColorBrush
        Public Property StatusFg As SolidColorBrush
    End Class

    ' ============================================================
    '  RENKLER
    ' ============================================================
    Private ReadOnly ColGreen As New SolidColorBrush(Color.FromRgb(0, 255, 136))
    Private ReadOnly ColCyan As New SolidColorBrush(Color.FromRgb(0, 212, 255))
    Private ReadOnly ColOrange As New SolidColorBrush(Color.FromRgb(255, 107, 53))
    Private ReadOnly ColRed As New SolidColorBrush(Color.FromRgb(255, 59, 92))
    Private ReadOnly ColPurple As New SolidColorBrush(Color.FromRgb(168, 85, 247))
    Private ReadOnly ColYellow As New SolidColorBrush(Color.FromRgb(255, 215, 0))
    Private ReadOnly ColGreenDim As New SolidColorBrush(Color.FromArgb(50, 0, 255, 136))
    Private ReadOnly ColOrangeDim As New SolidColorBrush(Color.FromArgb(50, 255, 107, 53))
    Private ReadOnly ColRedDim As New SolidColorBrush(Color.FromArgb(50, 255, 59, 92))
    Private ReadOnly ColCyanDim As New SolidColorBrush(Color.FromArgb(50, 0, 212, 255))
    Private ReadOnly ColTextMuted As New SolidColorBrush(Color.FromRgb(90, 106, 133))
    Private ReadOnly ColCardBg As New SolidColorBrush(Color.FromRgb(26, 34, 53))

    ' ============================================================
    '  GERCEK ZAMANLI OBSERVABLE KOLEKSIYONLAR
    ' ============================================================
    Private _headingList As New ObservableCollection(Of HeadingItem)()
    Private _linkList As New ObservableCollection(Of LinkItem)()
    Private _imageList As New ObservableCollection(Of ImageItem)()
    Private _checkList As New ObservableCollection(Of ChecklistItem)()
    Private _techList As New ObservableCollection(Of TechItem)()
    Private _keywordList As New ObservableCollection(Of KeywordItem)()
    Private _robotsList As New ObservableCollection(Of TechItem)()
    Private _compareList As New ObservableCollection(Of CompareItem)()
    Private _historyList As New ObservableCollection(Of HistoryItem)()
    Private _currentDoc As HtmlDocument = Nothing
    Private _currentUrl As String = String.Empty
    Private _currentScore As String = "--"
    Private _currentSpeedMs As String = "--"

    Public Sub New()
        InitializeComponent()
        ' Listeleri onceden bagla - veri eklenince otomatik guncellenir
        HeadingItems.ItemsSource = _headingList
        LinkItems.ItemsSource = _linkList
        ImageItems.ItemsSource = _imageList
        ChecklistItems.ItemsSource = _checkList
        TechItems.ItemsSource = _techList
        KeywordItems.ItemsSource = _keywordList
        RobotsItems.ItemsSource = _robotsList
        CompareItems.ItemsSource = _compareList
        HistoryItems.ItemsSource = _historyList
        LoadHistory()
        SendStartupPingAsync() ' Arka planda sessizce isteği atar ve unutur
    End Sub

    ' ============================================================
    '  UI OLAYLARI
    ' ============================================================
    Private Sub TxtUrl_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then BtnAnalyze_Click(sender, Nothing)
    End Sub

    ' ============================================================
    '  YENI BUTON HANDLER'LARI
    ' ============================================================
    Private Sub BtnExportCsv_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim sfd As New Microsoft.Win32.SaveFileDialog()
            sfd.Filter = "CSV Dosyasi (*.csv)|*.csv"
            sfd.FileName = "SEO_Links_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".csv"
            If sfd.ShowDialog() = True Then
                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine("Tip,URL,Anchor Text,HTTP Kodu")
                For Each item As LinkItem In _linkList
                    Dim line As String = item.LinkType & "," &
                        Chr(34) & item.RawUrl.Replace(Chr(34), "") & Chr(34) & "," &
                        Chr(34) & item.AnchorText.Replace(Chr(34), "") & Chr(34) & "," &
                        item.StatusCode
                    sb.AppendLine(line)
                Next
                System.IO.File.WriteAllText(sfd.FileName, sb.ToString(), System.Text.Encoding.UTF8)
                TxtStatus.Text = "CSV kaydedildi: " & sfd.FileName
            End If
        Catch ex As Exception
            MessageBox.Show("CSV hatasi: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnRakip_Click(sender As Object, e As RoutedEventArgs)
        ' Rakip sekmesine git (sekme indeksi: Genel=0, Baslik=1, Linkler=2, Gorseller=3, Teknik=4, Icerik=5, Rakip=6, Gecmis=7, Kaynak=8)
        MainTabs.SelectedIndex = 6
        TxtRakipUrl.Focus()
    End Sub

    Private Async Sub BtnRakipAnalyze_Click(sender As Object, e As RoutedEventArgs)
        Dim rakipUrl As String = TxtRakipUrl.Text.Trim()
        If String.IsNullOrEmpty(rakipUrl) OrElse rakipUrl = "https://" Then
            MessageBox.Show("Lutfen gecerli bir rakip URL girin.", "Uyari", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        If Not rakipUrl.StartsWith("http://") AndAlso Not rakipUrl.StartsWith("https://") Then
            rakipUrl = "https://" & rakipUrl
            TxtRakipUrl.Text = rakipUrl
        End If
        If _currentDoc Is Nothing Then
            MessageBox.Show("Once ana siteyi analiz edin.", "Uyari", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        RakipLoading.Visibility = Visibility.Visible
        BtnRakipAnalyze.IsEnabled = False
        _compareList.Clear()
        Try
            Await AnalyzeRakipAsync(rakipUrl)
        Catch ex As Exception
            MessageBox.Show("Rakip analiz hatasi: " & ex.Message)
        Finally
            RakipLoading.Visibility = Visibility.Collapsed
            BtnRakipAnalyze.IsEnabled = True
        End Try
    End Sub

    ' ============================================================
    '  GEÇMİŞİ TEMİZLE BUTONU
    ' ============================================================
    Private Sub BtnClearHistory_Click(sender As Object, e As RoutedEventArgs)
        ' 1. Arayüzdeki (RAM'deki) listeyi temizle
        _historyList.Clear()

        ' 2. Bilgisayara kaydettiğimiz .txt dosyasını fiziksel olarak sil
        Try
            If System.IO.File.Exists(HistoryFilePath) Then
                System.IO.File.Delete(HistoryFilePath)
            End If

            TxtStatus.Text = "Analiz geçmişi kalıcı olarak temizlendi."
            MessageBox.Show("Tüm analiz geçmişi başarıyla silindi.", "Temizlendi", MessageBoxButton.OK, MessageBoxImage.Information)
        Catch ex As Exception
            TxtStatus.Text = "Geçmiş temizlendi ancak dosya silinemedi."
            MessageBox.Show("Dosya silinirken bir hata oluştu: " & ex.Message, "Hata", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
    Private Sub HistoryItem_Click(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
        Dim border = TryCast(sender, System.Windows.Controls.Border)
        If border Is Nothing Then Return
        Dim item = TryCast(border.DataContext, HistoryItem)
        If item Is Nothing Then Return
        TxtUrl.Text = item.Url
        BtnAnalyze_Click(Nothing, Nothing)
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As RoutedEventArgs)
        TxtUrl.Text = "https://"
        ResetUI()
        TxtStatus.Text = "Temizlendi."
    End Sub

    Private Sub BtnCopyHtml_Click(sender As Object, e As RoutedEventArgs)
        If Not String.IsNullOrEmpty(TxtRawHtml.Text) Then
            Clipboard.SetText(TxtRawHtml.Text)
            TxtStatus.Text = "Kaynak kod panoya kopyalandi."
        End If
    End Sub

    Private Async Sub BtnAnalyze_Click(sender As Object, e As RoutedEventArgs)
        Dim url As String = TxtUrl.Text.Trim()
        If String.IsNullOrEmpty(url) OrElse url = "https://" Then
            MessageBox.Show("Lutfen gecerli bir URL girin.", "Uyari", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        If Not url.StartsWith("http://") AndAlso Not url.StartsWith("https://") Then
            url = "https://" & url
            TxtUrl.Text = url
        End If
        BtnAnalyze.IsEnabled = False
        LoadingPanel.Visibility = Visibility.Visible
        TxtStatus.Text = "Baglaniliyor..."
        ResetUI()
        Try
            Await AnalyzeAsync(url)
        Catch ex As Exception
            TxtStatus.Text = "Hata: " & ex.Message
            MessageBox.Show("Analiz hatasi:" & Chr(13) & ex.Message, "Hata", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            BtnAnalyze.IsEnabled = True
            LoadingPanel.Visibility = Visibility.Collapsed
        End Try
    End Sub

    Private Async Function AnalyzeAsync(url As String) As Task
        ' ===== 1. GÜVENLİK ZIRHI: Hata yakalama bloğu başlatılıyor =====
        Try
            ' --- 1. HTTP ISTEGI ---
            UpdateStatus("HTTP istegi gonderiliyor...")
            Dim handler As New HttpClientHandler() With {.AllowAutoRedirect = True}

            ' HttpClient'i Using bloğuna aldık (İş bitince RAM'den otomatik silinir)
            Dim html As String = ""
            Dim ms As Long = 0

            Using client As New HttpClient(handler)
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36")
                client.Timeout = TimeSpan.FromSeconds(20)

                Dim sw As New System.Diagnostics.Stopwatch()
                sw.Start() ' Kronometreyi TAM BURADA başlatıyoruz ki ayar süreleri hıza yansımasın

                html = Await client.GetStringAsync(url)

                sw.Stop()
                ms = sw.ElapsedMilliseconds
            End Using

            ' Sayfa hizini hemen goster
            TxtSpeed.Text = ms.ToString() & " ms"
            If ms < 800 Then
                TxtSpeed.Foreground = ColGreen
                TxtSpeedNote.Text = "Hizli yanit (iyi)"
            ElseIf ms < 2000 Then
                TxtSpeed.Foreground = ColYellow
                TxtSpeedNote.Text = "Ortalama yanit"
            Else
                TxtSpeed.Foreground = ColRed
                TxtSpeedNote.Text = "Yavas yanit (optimize edin)"
            End If

            ' --- 2. HTML PARSE ---
            UpdateStatus("HTML ayristiriliyor...")
            Dim doc As New HtmlDocument()
            doc.LoadHtml(html)

            ' Kaynak kodu arka planda yukle (UI donmasin)
            _currentDoc = doc
            _currentUrl = url
            Await Application.Current.Dispatcher.InvokeAsync(Sub()
                                                                 TxtRawHtml.Text = html
                                                             End Sub, Threading.DispatcherPriority.Background)

            ' --- 3. TITLE + META ---
            UpdateStatus("Sayfa basligi ve meta veriler analiz ediliyor...")
            AnalyzeTitle(doc)
            AnalyzeDescription(doc)
            Await Task.Yield() ' UI'nin nefes almasina izin ver

            ' --- 4. BASLIKLAR ---
            UpdateStatus("Baslik yapisi analiz ediliyor...")
            AnalyzeHeadings(doc)
            Await Task.Yield()

            ' --- 5. GORSELLER ---
            UpdateStatus("Gorsel alt etiketleri kontrol ediliyor...")
            AnalyzeImages(doc)
            Await Task.Yield()

            ' --- 6. TEKNIK SEO ---
            UpdateStatus("Teknik SEO parametreleri kontrol ediliyor...")
            AnalyzeTechnical(doc, url)
            Await Task.Yield()

            ' --- 7. KELIME + ANAHTAR ---
            UpdateStatus("Icerik ve anahtar kelime analizi yapiliyor...")
            AnalyzeWordCount(doc)
            AnalyzeKeywords(doc)
            Await Task.Yield()

            ' --- 8. KONTROL LISTESI + SKOR ---
            UpdateStatus("SEO skoru hesaplaniyor...")
            BuildChecklist(doc, url)
            CalculateScore(doc, url, ms)
            Await Task.Yield()

            ' --- 9. LINK LISTESI ---
            UpdateStatus("Linkler listeleniyor...")
            Await AnalyzeLinksAsync(doc, url)

            ' --- 10. ICERIK KALITESI + ROBOTS/SITEMAP ---
            UpdateStatus("Icerik kalitesi ve robots/sitemap kontrol ediliyor...")
            AnalyzeReadability(doc)
            AnalyzeTextHtmlRatio(doc, html) ' (Not: Bu metodu ayrı yazdıysan sorun yok)
            Await AnalyzeRobotsAndSitemapAsync(url)

            ' --- 11. GECMISE EKLE ---
            _currentScore = TxtScore.Text
            _currentSpeedMs = ms.ToString() & " ms"
            AddToHistory(url, ms)

            BtnExport.IsEnabled = True
            ' BtnExportCsv isimli butonun varsa aktif olur, yoksa silebilirsin.
            If BtnExportCsv IsNot Nothing Then BtnExportCsv.IsEnabled = True

            TxtStatus.Text = "Analiz tamamlandi: " & url & "  (" & ms.ToString() & " ms)"
            LoadingPanel.Visibility = Visibility.Collapsed ' Başlangıçtaki Loading ekranını kapat

        Catch ex As HttpRequestException
            MessageBox.Show("Siteye ulaşılamadı. Lütfen URL'yi (https:// ile) doğru yazdığınızdan ve sitenin açık olduğundan emin olun." & vbCrLf & vbCrLf & "Hata Detayı: " & ex.Message, "Bağlantı Hatası", MessageBoxButton.OK, MessageBoxImage.Warning)
            TxtStatus.Text = "Bağlantı hatası: Analiz iptal edildi."
            LoadingPanel.Visibility = Visibility.Collapsed

        Catch ex As TaskCanceledException
            MessageBox.Show("Site 20 saniye içinde yanıt vermedi (Zaman aşımı).", "Zaman Aşımı", MessageBoxButton.OK, MessageBoxImage.Warning)
            TxtStatus.Text = "Zaman aşımı: Analiz iptal edildi."
            LoadingPanel.Visibility = Visibility.Collapsed

        Catch ex As Exception
            MessageBox.Show("Analiz sırasında beklenmeyen bir hata oluştu:" & vbCrLf & ex.Message, "Hata", MessageBoxButton.OK, MessageBoxImage.Error)
            TxtStatus.Text = "Kritik hata: " & ex.Message
            LoadingPanel.Visibility = Visibility.Collapsed
        End Try
    End Function

    ' Dispatcher'a gerek kalmadan status guncelle
    Private Sub UpdateStatus(msg As String)
        TxtLoadingDetail.Text = msg
        TxtStatus.Text = msg
    End Sub

    ' ============================================================
    '  LİNKLER VE KIRIK LİNK (404) TARAYICI (GÜNCELLENDİ)
    ' ============================================================
    Private Async Function AnalyzeLinksAsync(doc As HtmlDocument, baseUrl As String) As Task
        Dim nodes = doc.DocumentNode.SelectNodes("//a[@href]")

        ' Toplam ham link sayısını baştan alıyoruz (Hızlı istatistikler için)
        Dim rawLinkCount As Integer = If(nodes IsNot Nothing, nodes.Count, 0)

        If nodes Is Nothing OrElse rawLinkCount = 0 Then
            StatLinks.Text = "0"
            Return
        End If

        Dim baseUri As Uri = Nothing
        Uri.TryCreate(baseUrl, UriKind.Absolute, baseUri)
        Dim internalCount As Integer = 0
        Dim externalCount As Integer = 0
        Dim seen As New HashSet(Of String)()

        ' Önceki analizden kalan listeyi temizle
        _linkList.Clear()

        ' Önce tüm benzersiz linkleri listele (Durumları "..." olarak başlar)
        For Each n As HtmlNode In nodes
            Dim href As String = n.GetAttributeValue("href", "")
            If String.IsNullOrEmpty(href) OrElse href.StartsWith("#") OrElse
               href.StartsWith("javascript") OrElse href.StartsWith("mailto:") OrElse
               href.StartsWith("tel:") Then Continue For

            Dim anchor As String = WebUtility.HtmlDecode(n.InnerText.Trim())
            Dim isInternal As Boolean = False
            Dim absUrl As String = href

            If baseUri IsNot Nothing Then
                Dim abs As Uri = Nothing
                If Uri.TryCreate(baseUri, href, abs) Then
                    isInternal = (abs.Host = baseUri.Host)
                    absUrl = abs.ToString()
                End If
            End If

            ' Aynı URL varsa atla (Tekilleştirme)
            If seen.Contains(absUrl) Then Continue For
            seen.Add(absUrl)

            If isInternal Then internalCount += 1 Else externalCount += 1

            Dim displayUrl As String = If(absUrl.Length > 90, absUrl.Substring(0, 90) & "...", absUrl)
            Dim displayAnchor As String = If(String.IsNullOrEmpty(anchor), "(anchor yok)",
                                             If(anchor.Length > 35, anchor.Substring(0, 35) & "...", anchor))

            ' Listeye ekle - ObservableCollection anında arayüze yansıtır
            _linkList.Add(New LinkItem With {
                .Url = displayUrl,
                .RawUrl = absUrl,
                .AnchorText = displayAnchor,
                .LinkType = If(isInternal, "İÇ", "DIŞ"),
                .TypeColor = If(isInternal, ColCyanDim, ColOrangeDim),
                .StatusCode = "...",
                .StatusColor = ColTextMuted
            })
        Next

        ' Ham link sayısını ve benzersiz link istatistiklerini arayüze bas
        StatLinks.Text = rawLinkCount.ToString()
        TxtInternalCount.Text = internalCount.ToString()
        TxtExternalCount.Text = externalCount.ToString()

        ' --- 404 TARAMASI BAŞLIYOR ---
        UpdateStatus("Bulunan " & _linkList.Count.ToString() & " benzersiz link test ediliyor...")

        Dim brokenCount As Integer = 0
        Dim httpHandler As New HttpClientHandler() With {.AllowAutoRedirect = True}

        Using httpClient As New HttpClient(httpHandler)
            httpClient.Timeout = TimeSpan.FromSeconds(8)
            httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

            Dim semaphore As New System.Threading.SemaphoreSlim(15)
            Dim tasks As New List(Of Task)()

            For i As Integer = 0 To _linkList.Count - 1
                Dim currentItem As LinkItem = _linkList(i)

                tasks.Add(Task.Run(Async Function()
                                       Await semaphore.WaitAsync()
                                       Dim resultCode As String = "ERR"
                                       Dim resultColor As SolidColorBrush = ColRed
                                       Dim isBroken As Boolean = False

                                       Try
                                           ' Performans için HEAD isteği
                                           Dim request As New HttpRequestMessage(HttpMethod.Head, currentItem.RawUrl)
                                           Dim response As HttpResponseMessage = Await httpClient.SendAsync(request)

                                           ' HEAD reddedilirse GET ile tekrar dene
                                           If response.StatusCode = HttpStatusCode.MethodNotAllowed Then
                                               request = New HttpRequestMessage(HttpMethod.Get, currentItem.RawUrl)
                                               response = Await httpClient.SendAsync(request)
                                           End If

                                           Dim code As Integer = CInt(response.StatusCode)
                                           If response.IsSuccessStatusCode Then
                                               resultCode = code.ToString()
                                               resultColor = ColGreen
                                           ElseIf code >= 300 AndAlso code < 400 Then
                                               resultCode = code.ToString() & " R" ' Yönlendirme
                                               resultColor = ColYellow
                                           Else
                                               resultCode = code.ToString()
                                               resultColor = ColRed
                                               isBroken = True
                                           End If
                                       Catch ex As Exception
                                           resultCode = "ERR"
                                           resultColor = ColRed
                                           isBroken = True
                                       Finally
                                           semaphore.Release()
                                       End Try

                                       ' Kırık link sayacını güvenli (Thread-Safe) artır
                                       If isBroken Then
                                           System.Threading.Interlocked.Increment(brokenCount)
                                       End If

                                       ' UI Güncellemesi (Tıkır tıkır canlı güncellenir)
                                       Application.Current.Dispatcher.Invoke(Sub()
                                                                                 currentItem.StatusCode = resultCode
                                                                                 currentItem.StatusColor = resultColor

                                                                                 If isBroken Then
                                                                                     TxtBrokenCount.Text = brokenCount.ToString()
                                                                                 End If

                                                                                 ' ObservableCollection'ı tetikle ki ekrandaki satır güncellensin
                                                                                 Dim itemIdx As Integer = _linkList.IndexOf(currentItem)
                                                                                 If itemIdx >= 0 Then
                                                                                     _linkList(itemIdx) = currentItem
                                                                                 End If
                                                                             End Sub)
                                   End Function))
            Next

            Await Task.WhenAll(tasks)
        End Using

        ' Finalde her ihtimale karşı kırık link sayacını son kez eşitle
        TxtBrokenCount.Text = brokenCount.ToString()
    End Function
    ' ============================================================
    '  AÇILIŞ PİNGİ (Sessiz Ziyaret)
    ' ============================================================
    Private Async Sub SendStartupPingAsync()
        Try
            Using client As New HttpClient()
                ' Programın normal bir tarayıcı gibi görünmesi için User-Agent ekliyoruz
                client.DefaultRequestHeaders.Add("User-Agent", "SEOAnalyzerPro/1.0 (Windows Desktop)")

                ' Sadece başlıkları okuyan (gövdeyi indirmeyen) çok hafif bir istek atıyoruz
                Dim request As New HttpRequestMessage(HttpMethod.Head, "https://bilgikasabasi.com/")
                Await client.SendAsync(request)
            End Using
        Catch ex As Exception
            ' Eğer internet yoksa veya site kapalıysa programın çökmemesi için hatayı sessizce yutuyoruz
        End Try
    End Sub
    ' ============================================================
    '  TITLE
    ' ============================================================
    Private Sub AnalyzeTitle(doc As HtmlDocument)
        Dim node = doc.DocumentNode.SelectSingleNode("//title")
        Dim title As String = If(node IsNot Nothing, node.InnerText.Trim(), "")
        TxtTitle.Text = If(String.IsNullOrEmpty(title), "(Baslik bulunamadi)", title)
        Dim len As Integer = title.Length
        TxtTitleLen.Text = len.ToString() & " karakter"
        If len = 0 Then
            TitleProgress.Value = 0
            TitleProgress.Foreground = ColRed
            SetBadge(TitleStatusBadge, TitleStatus, "EKSIK", ColRedDim, ColRed)
        ElseIf len >= 50 AndAlso len <= 60 Then
            TitleProgress.Value = 100
            TitleProgress.Foreground = ColGreen
            SetBadge(TitleStatusBadge, TitleStatus, "IDEAL", ColGreenDim, ColGreen)
        ElseIf len < 50 Then
            TitleProgress.Value = CInt(len / 50.0 * 80)
            TitleProgress.Foreground = ColOrange
            SetBadge(TitleStatusBadge, TitleStatus, "KISA", ColOrangeDim, ColOrange)
        Else
            TitleProgress.Value = 70
            TitleProgress.Foreground = ColOrange
            SetBadge(TitleStatusBadge, TitleStatus, "UZUN", ColOrangeDim, ColOrange)
        End If
    End Sub

    ' ============================================================
    '  DESCRIPTION
    ' ============================================================
    Private Sub AnalyzeDescription(doc As HtmlDocument)
        Dim node = doc.DocumentNode.SelectSingleNode("//meta[@name='description']")
        Dim desc As String = If(node IsNot Nothing, node.GetAttributeValue("content", ""), "")
        TxtDesc.Text = If(String.IsNullOrEmpty(desc), "(Meta aciklama bulunamadi)", desc)
        Dim len As Integer = desc.Length
        TxtDescLen.Text = len.ToString() & " karakter"
        If len = 0 Then
            DescProgress.Value = 0
            DescProgress.Foreground = ColRed
            SetBadge(DescStatusBadge, DescStatus, "EKSIK", ColRedDim, ColRed)
        ElseIf len >= 150 AndAlso len <= 160 Then
            DescProgress.Value = 100
            DescProgress.Foreground = ColGreen
            SetBadge(DescStatusBadge, DescStatus, "IDEAL", ColGreenDim, ColGreen)
        ElseIf len < 150 Then
            DescProgress.Value = CInt(len / 150.0 * 80)
            DescProgress.Foreground = ColOrange
            SetBadge(DescStatusBadge, DescStatus, "KISA", ColOrangeDim, ColOrange)
        Else
            DescProgress.Value = 70
            DescProgress.Foreground = ColOrange
            SetBadge(DescStatusBadge, DescStatus, "UZUN", ColOrangeDim, ColOrange)
        End If
    End Sub

    ' ============================================================
    '  BASLIKLAR - GERCEK ZAMANLI (H icin renk kadar ekle)
    ' ============================================================
    Private Sub AnalyzeHeadings(doc As HtmlDocument)
        Dim h1Count As Integer = 0
        Dim h2Count As Integer = 0
        For Each tag As String In New String() {"h1", "h2", "h3", "h4", "h5", "h6"}
            Dim nodes = doc.DocumentNode.SelectNodes("//" & tag)
            If nodes Is Nothing Then Continue For
            Dim color As SolidColorBrush
            Select Case tag
                Case "h1" : color = ColCyan : h1Count += nodes.Count
                Case "h2" : color = ColGreen : h2Count += nodes.Count
                Case "h3" : color = ColOrange
                Case "h4" : color = ColPurple
                Case Else : color = ColTextMuted
            End Select
            For Each n As HtmlNode In nodes
                Dim txt As String = WebUtility.HtmlDecode(n.InnerText.Trim())
                If Not String.IsNullOrWhiteSpace(txt) Then
                    _headingList.Add(New HeadingItem With {
                        .Tag = tag.ToUpper(),
                        .Text = If(txt.Length > 130, txt.Substring(0, 130) & "...", txt),
                        .TagColor = color
                    })
                End If
            Next
        Next
        StatH1.Text = h1Count.ToString()
        StatH2.Text = h2Count.ToString()
    End Sub

    ' ============================================================
    '  GORSELLER - GERCEK ZAMANLI
    ' ============================================================
    Private Sub AnalyzeImages(doc As HtmlDocument)
        Dim nodes = doc.DocumentNode.SelectNodes("//img")
        If nodes Is Nothing Then
            StatImages.Text = "0"
            TxtAltOkCount.Text = "0"
            TxtAltMissingCount.Text = "0"
            Return
        End If
        Dim altOk As Integer = 0
        Dim altMissing As Integer = 0
        For Each n As HtmlNode In nodes
            ' Lazy loading destegi: src > data-src > data-lazy-src > data-original sirasiyla dene
            Dim src As String = ""
            Dim srcLabel As String = ""
            Dim candidates() As String = {"src", "data-src", "data-lazy-src", "data-original", "data-lazy", "data-srcset"}
            For Each attr As String In candidates
                Dim val As String = n.GetAttributeValue(attr, "")
                If Not String.IsNullOrWhiteSpace(val) AndAlso Not val.StartsWith("data:image") Then
                    src = val
                    srcLabel = If(attr = "src", "", "[" & attr & "] ")
                    Exit For
                End If
            Next
            If String.IsNullOrEmpty(src) Then src = "(src bulunamadi)" : srcLabel = ""

            Dim alt As String = n.GetAttributeValue("alt", "")
            Dim hasAlt As Boolean = Not String.IsNullOrWhiteSpace(alt)
            If hasAlt Then altOk += 1 Else altMissing += 1

            Dim displaySrc As String = srcLabel & If(src.Length > 80, src.Substring(0, 80) & "...", src)
            _imageList.Add(New ImageItem With {
                .Src = displaySrc,
                .AltText = If(hasAlt, "Alt: " & If(alt.Length > 60, alt.Substring(0, 60) & "...", alt), "(alt etiketi eksik)"),
                .AltStatus = If(hasAlt, "ALT VAR", "ALT YOK"),
                .AltColor = If(hasAlt, ColGreenDim, ColRedDim),
                .AltBadgeColor = If(hasAlt, ColGreenDim, ColRedDim)
            })
        Next
        StatImages.Text = _imageList.Count.ToString()
        TxtAltOkCount.Text = altOk.ToString()
        TxtAltMissingCount.Text = altMissing.ToString()
    End Sub

    ' ============================================================
    '  KELIME SAYISI
    ' ============================================================
    Private Sub AnalyzeWordCount(doc As HtmlDocument)
        Dim bodyNode = doc.DocumentNode.SelectSingleNode("//body")
        If bodyNode Is Nothing Then
            StatWords.Text = "0"
            Return
        End If
        Dim cleanText As String = GetCleanText(bodyNode)
        ' \p{L} = tum Unicode harfler (Turkce dahil), \s = bosluk - sayi/sembol/noktalama kaldir
        Dim rgx As New System.Text.RegularExpressions.Regex("[^\p{L}\s]")
        cleanText = rgx.Replace(cleanText, " ")
        Dim separators() As Char = {" "c, Chr(13), Chr(10), Chr(9), Chr(160)}
        Dim words() As String = cleanText.Split(separators, StringSplitOptions.RemoveEmptyEntries)
        Dim realWords As Integer = 0
        For Each w As String In words
            If w.Trim().Length > 2 Then realWords += 1
        Next
        StatWords.Text = realWords.ToString("N0")
    End Sub

    ' ============================================================
    '  ANAHTAR KELIME YOGUNLUGU
    ' ============================================================
    Private Sub AnalyzeKeywords(doc As HtmlDocument)
        Dim bodyNode = doc.DocumentNode.SelectSingleNode("//body")
        If bodyNode Is Nothing Then Return
        Dim cleanText As String = GetCleanText(bodyNode).ToLower()
        ' \p{L} = tum Unicode harfler (Turkce dahil)
        Dim rgx As New System.Text.RegularExpressions.Regex("[^\p{L}\s]")
        cleanText = rgx.Replace(cleanText, " ")
        Dim separators() As Char = {" "c, Chr(13), Chr(10), Chr(9)}
        Dim words() As String = cleanText.Split(separators, StringSplitOptions.RemoveEmptyEntries)
        Dim stopWords As New HashSet(Of String)({"bir", "bu", "ve", "ile", "de", "da", "ki", "mi", "mu",
            "the", "and", "for", "are", "was", "but", "not", "you", "all", "can",
            "her", "icin", "olan", "olarak", "daha", "ise", "den", "dan", "biz",
            "ben", "sen", "bunu", "buna", "bunun", "yani", "ama", "ya", "ne",
            "at", "to", "of", "is", "it", "be", "as", "or", "an", "its", "by",
            "that", "this", "with", "from", "they", "will", "have", "has", "had",
            "iki", "uc", "dort", "bes", "alti", "yedi", "sekiz", "dokuz",
            "olup", "gibi", "kadar", "sonra", "once", "gore", "icin",
            "en", "cok", "az", "var", "yok", "hem", "nasil", "neden"})
        Dim freq As New Dictionary(Of String, Integer)()
        For Each w As String In words
            If w.Length < 4 Then Continue For
            If stopWords.Contains(w) Then Continue For
            If freq.ContainsKey(w) Then
                freq(w) += 1
            Else
                freq(w) = 1
            End If
        Next
        Dim sorted As New List(Of KeyValuePair(Of String, Integer))(freq)
        sorted.Sort(Function(a, b) b.Value.CompareTo(a.Value))
        Dim maxCount As Integer = If(sorted.Count > 0, sorted(0).Value, 1)
        Dim top As Integer = Math.Min(8, sorted.Count)
        For i As Integer = 0 To top - 1
            Dim kv = sorted(i)
            _keywordList.Add(New KeywordItem With {
                .Word = kv.Key,
                .CountText = kv.Value.ToString() & "x",
                .Percent = CDbl(kv.Value) / CDbl(maxCount) * 100.0
            })
        Next
    End Sub

    ' ============================================================
    '  TEKNIK SEO - GERCEK ZAMANLI
    ' ============================================================
    Private Sub AnalyzeTechnical(doc As HtmlDocument, url As String)
        ' Lang
        Dim htmlNode = doc.DocumentNode.SelectSingleNode("//html")
        Dim lang As String = If(htmlNode IsNot Nothing, htmlNode.GetAttributeValue("lang", ""), "")
        _techList.Add(MakeTechItem("Dil (lang)", If(String.IsNullOrEmpty(lang), "Tanimlanmamis", lang),
            If(String.IsNullOrEmpty(lang), "EKSIK", "OK"),
            If(String.IsNullOrEmpty(lang), ColRedDim, ColGreenDim),
            If(String.IsNullOrEmpty(lang), ColRed, ColGreen)))

        ' Charset
        Dim charsetNode = doc.DocumentNode.SelectSingleNode("//meta[@charset]")
        Dim hasCharset As Boolean = (charsetNode IsNot Nothing OrElse
            doc.DocumentNode.SelectSingleNode("//meta[@http-equiv='Content-Type']") IsNot Nothing)
        Dim charsetVal As String = If(charsetNode IsNot Nothing, charsetNode.GetAttributeValue("charset", ""), "")
        _techList.Add(MakeTechItem("Karakter Seti", If(String.IsNullOrEmpty(charsetVal), If(hasCharset, "Mevcut", "Bulunamadi"), charsetVal),
            If(hasCharset, "OK", "EKSIK"),
            If(hasCharset, ColGreenDim, ColRedDim),
            If(hasCharset, ColGreen, ColRed)))

        ' Viewport
        Dim vpNode = doc.DocumentNode.SelectSingleNode("//meta[@name='viewport']")
        Dim vpVal As String = If(vpNode IsNot Nothing, vpNode.GetAttributeValue("content", ""), "")
        _techList.Add(MakeTechItem("Viewport", If(String.IsNullOrEmpty(vpVal), "Bulunamadi", vpVal),
            If(vpNode IsNot Nothing, "OK", "EKSIK"),
            If(vpNode IsNot Nothing, ColGreenDim, ColRedDim),
            If(vpNode IsNot Nothing, ColGreen, ColRed)))

        ' Canonical
        Dim canonNode = doc.DocumentNode.SelectSingleNode("//link[@rel='canonical']")
        Dim canonVal As String = If(canonNode IsNot Nothing, canonNode.GetAttributeValue("href", ""), "")
        _techList.Add(MakeTechItem("Canonical URL", If(String.IsNullOrEmpty(canonVal), "Bulunamadi", canonVal),
            If(canonNode IsNot Nothing, "OK", "EKSIK"),
            If(canonNode IsNot Nothing, ColGreenDim, ColOrangeDim),
            If(canonNode IsNot Nothing, ColGreen, ColOrange)))

        ' Meta Keywords
        Dim kwNode = doc.DocumentNode.SelectSingleNode("//meta[@name='keywords']")
        Dim kwVal As String = If(kwNode IsNot Nothing, kwNode.GetAttributeValue("content", ""), "")
        _techList.Add(MakeTechItem("Meta Keywords", If(String.IsNullOrEmpty(kwVal), "Tanimlanmamis", kwVal),
            If(kwNode IsNot Nothing, "VAR", "YOK"),
            If(kwNode IsNot Nothing, ColCyanDim, ColOrangeDim),
            If(kwNode IsNot Nothing, ColCyan, ColOrange)))

        ' OG Tags
        Dim ogTitle = doc.DocumentNode.SelectSingleNode("//meta[@property='og:title']")
        Dim ogDesc = doc.DocumentNode.SelectSingleNode("//meta[@property='og:description']")
        Dim ogImg = doc.DocumentNode.SelectSingleNode("//meta[@property='og:image']")
        Dim ogCount As Integer = 0
        If ogTitle IsNot Nothing Then ogCount += 1
        If ogDesc IsNot Nothing Then ogCount += 1
        If ogImg IsNot Nothing Then ogCount += 1
        _techList.Add(MakeTechItem("Open Graph Tags", ogCount.ToString() & "/3 etiket (title, desc, image)",
            If(ogCount = 3, "TAMAM", If(ogCount > 0, "KISMI", "EKSIK")),
            If(ogCount = 3, ColGreenDim, If(ogCount > 0, ColOrangeDim, ColRedDim)),
            If(ogCount = 3, ColGreen, If(ogCount > 0, ColOrange, ColRed))))

        ' Twitter Card
        Dim twCard = doc.DocumentNode.SelectSingleNode("//meta[@name='twitter:card']")
        _techList.Add(MakeTechItem("Twitter Card", If(twCard IsNot Nothing, twCard.GetAttributeValue("content", "Mevcut"), "Bulunamadi"),
            If(twCard IsNot Nothing, "OK", "YOK"),
            If(twCard IsNot Nothing, ColGreenDim, ColOrangeDim),
            If(twCard IsNot Nothing, ColGreen, ColOrange)))

        ' Robots Meta
        Dim robotsNode = doc.DocumentNode.SelectSingleNode("//meta[@name='robots']")
        Dim robotsVal As String = If(robotsNode IsNot Nothing, robotsNode.GetAttributeValue("content", ""), "index, follow (varsayilan)")
        Dim isNoindex As Boolean = robotsVal.ToLower().Contains("noindex")
        _techList.Add(MakeTechItem("Robots Meta", robotsVal,
            If(isNoindex, "NOINDEX", "OK"),
            If(isNoindex, ColRedDim, ColGreenDim),
            If(isNoindex, ColRed, ColGreen)))

        ' Schema.org
        Dim schemaNodes = doc.DocumentNode.SelectNodes("//script[@type='application/ld+json']")
        Dim schemaCount As Integer = If(schemaNodes IsNot Nothing, schemaNodes.Count, 0)
        _techList.Add(MakeTechItem("Schema.org (JSON-LD)", If(schemaCount = 0, "Bulunamadi", schemaCount.ToString() & " yapisal veri blogu"),
            If(schemaCount > 0, "OK", "EKSIK"),
            If(schemaCount > 0, ColGreenDim, ColOrangeDim),
            If(schemaCount > 0, ColGreen, ColOrange)))

        ' HTTPS
        Dim isHttps As Boolean = url.StartsWith("https://")
        _techList.Add(MakeTechItem("Protokol", If(isHttps, "HTTPS - Guvenli baglanti", "HTTP - Guvensiz!"),
            If(isHttps, "HTTPS", "HTTP"),
            If(isHttps, ColGreenDim, ColRedDim),
            If(isHttps, ColGreen, ColRed)))

        ' HTML Boyutu
        Dim sizeKb As Integer = doc.DocumentNode.OuterHtml.Length \ 1024
        _techList.Add(MakeTechItem("HTML Boyutu", sizeKb.ToString() & " KB",
            If(sizeKb < 100, "IYI", "BUYUK"),
            If(sizeKb < 100, ColGreenDim, ColOrangeDim),
            If(sizeKb < 100, ColGreen, ColOrange)))

        ' Favicon
        Dim favicon = doc.DocumentNode.SelectSingleNode("//link[@rel='icon' or @rel='shortcut icon']")
        _techList.Add(MakeTechItem("Favicon", If(favicon IsNot Nothing, favicon.GetAttributeValue("href", "Mevcut"), "Bulunamadi"),
            If(favicon IsNot Nothing, "OK", "YOK"),
            If(favicon IsNot Nothing, ColGreenDim, ColOrangeDim),
            If(favicon IsNot Nothing, ColGreen, ColOrange)))

        ' Inline CSS miktari
        Dim inlineStyles = doc.DocumentNode.SelectNodes("//*[@style]")
        Dim inlineCount As Integer = If(inlineStyles IsNot Nothing, inlineStyles.Count, 0)
        _techList.Add(MakeTechItem("Inline Style Sayisi", inlineCount.ToString() & " element",
            If(inlineCount < 10, "IYI", If(inlineCount < 30, "ORTA", "COK")),
            If(inlineCount < 10, ColGreenDim, If(inlineCount < 30, ColOrangeDim, ColRedDim)),
            If(inlineCount < 10, ColGreen, If(inlineCount < 30, ColOrange, ColRed))))
    End Sub

    Private Function MakeTechItem(label As String, value As String, status As String,
                                   bg As SolidColorBrush, fg As SolidColorBrush) As TechItem
        Return New TechItem With {
            .Label = label, .Value = value, .Status = status,
            .BadgeColor = bg, .StatusFg = fg
        }
    End Function

    ' ============================================================
    '  KONTROL LISTESI - GERCEK ZAMANLI
    ' ============================================================
    Private Sub BuildChecklist(doc As HtmlDocument, url As String)
        ' 1. Title
        Dim titleNode = doc.DocumentNode.SelectSingleNode("//title")
        Dim titleLen As Integer = If(titleNode IsNot Nothing, titleNode.InnerText.Trim().Length, 0)
        Dim titleOk As Boolean = (titleLen >= 50 AndAlso titleLen <= 60)
        _checkList.Add(NewCheck("Sayfa Basligi",
            If(titleLen = 0, "Baslik bulunamadi - kritik eksiklik",
               If(titleOk, titleLen.ToString() & " karakter - ideal aralikta",
                  titleLen.ToString() & " karakter - ideal: 50-60")),
            If(titleLen = 0, "X", If(titleOk, "OK", "!")),
            If(titleLen = 0, ColRed, If(titleOk, ColGreen, ColOrange))))

        ' 2. Meta Desc
        Dim descNode = doc.DocumentNode.SelectSingleNode("//meta[@name='description']")
        Dim descLen As Integer = If(descNode IsNot Nothing, descNode.GetAttributeValue("content", "").Length, 0)
        Dim descOk As Boolean = (descLen >= 150 AndAlso descLen <= 160)
        _checkList.Add(NewCheck("Meta Aciklama",
            If(descLen = 0, "Meta description eksik",
               If(descOk, descLen.ToString() & " karakter - ideal",
                  descLen.ToString() & " karakter - ideal: 150-160")),
            If(descLen = 0, "X", If(descOk, "OK", "!")),
            If(descLen = 0, ColRed, If(descOk, ColGreen, ColOrange))))

        ' 3. H1
        Dim h1Nodes = doc.DocumentNode.SelectNodes("//h1")
        Dim h1Count As Integer = If(h1Nodes IsNot Nothing, h1Nodes.Count, 0)
        _checkList.Add(NewCheck("H1 Baslik",
            If(h1Count = 0, "H1 etiketi yok", If(h1Count = 1, "Tek H1 - ideal", h1Count.ToString() & " H1 - fazla")),
            If(h1Count = 1, "OK", If(h1Count = 0, "X", "!")),
            If(h1Count = 1, ColGreen, If(h1Count = 0, ColRed, ColOrange))))

        ' 4. H2/H3
        Dim h2n = doc.DocumentNode.SelectNodes("//h2")
        Dim h3n = doc.DocumentNode.SelectNodes("//h3")
        Dim subH As Integer = If(h2n IsNot Nothing, h2n.Count, 0) + If(h3n IsNot Nothing, h3n.Count, 0)
        _checkList.Add(NewCheck("Icerik Hiyerarsisi (H2, H3)",
            If(subH > 0, subH.ToString() & " adet alt baslik bulundu", "Hic alt baslik yok - okunabilirlik zayif"),
            If(subH > 0, "OK", "!"),
            If(subH > 0, ColGreen, ColOrange)))

        ' 5. Linkler
        Dim linkNodes = doc.DocumentNode.SelectNodes("//a[@href]")
        Dim linksTotal As Integer = If(linkNodes IsNot Nothing, linkNodes.Count, 0)
        _checkList.Add(NewCheck("Sayfa Ici Linkler",
            If(linksTotal > 0, "Sayfada " & linksTotal.ToString() & " adet link mevcut", "Sayfada hic link yok"),
            If(linksTotal > 0, "OK", "!"),
            If(linksTotal > 0, ColGreen, ColOrange)))

        ' 6. HTTPS
        Dim isHttps As Boolean = url.StartsWith("https://")
        _checkList.Add(NewCheck("HTTPS / SSL",
            If(isHttps, "Guvenli HTTPS baglantisi aktif", "HTTP kullaniyor - SEO'ya zarar verir"),
            If(isHttps, "OK", "X"), If(isHttps, ColGreen, ColRed)))

        ' 7. Viewport
        Dim hasVp As Boolean = (doc.DocumentNode.SelectSingleNode("//meta[@name='viewport']") IsNot Nothing)
        _checkList.Add(NewCheck("Viewport (Mobil)",
            If(hasVp, "Mobil uyumlu viewport mevcut", "Viewport eksik - mobil uyumsuz"),
            If(hasVp, "OK", "X"), If(hasVp, ColGreen, ColRed)))

        ' 8. Canonical
        Dim hasCanon As Boolean = (doc.DocumentNode.SelectSingleNode("//link[@rel='canonical']") IsNot Nothing)
        _checkList.Add(NewCheck("Canonical Tag",
            If(hasCanon, "Canonical URL tanimlanmis", "Canonical eksik - duplicate content riski"),
            If(hasCanon, "OK", "!"), If(hasCanon, ColGreen, ColOrange)))

        ' 9. OG
        Dim hasOg As Boolean = (doc.DocumentNode.SelectSingleNode("//meta[@property='og:title']") IsNot Nothing)
        _checkList.Add(NewCheck("Open Graph",
            If(hasOg, "Sosyal medya etiketleri mevcut", "OG etiketleri eksik - sosyal paylasimlari etkiler"),
            If(hasOg, "OK", "!"), If(hasOg, ColGreen, ColOrange)))

        ' 10. Schema
        Dim hasSd As Boolean = (doc.DocumentNode.SelectSingleNode("//script[@type='application/ld+json']") IsNot Nothing)
        _checkList.Add(NewCheck("Schema.org",
            If(hasSd, "Yapisal veri (JSON-LD) mevcut", "Yapisal veri yok - rich snippet kaybi"),
            If(hasSd, "OK", "!"), If(hasSd, ColGreen, ColOrange)))

        ' 11. Alt Tags
        Dim imgNodes2 = doc.DocumentNode.SelectNodes("//img")
        Dim imgTotal As Integer = If(imgNodes2 IsNot Nothing, imgNodes2.Count, 0)
        Dim imgMissing As Integer = 0
        If imgNodes2 IsNot Nothing Then
            For Each imgN As HtmlNode In imgNodes2
                If String.IsNullOrWhiteSpace(imgN.GetAttributeValue("alt", "")) Then imgMissing += 1
            Next
        End If
        _checkList.Add(NewCheck("Gorsel Alt Etiketleri",
            If(imgTotal = 0, "Gorsel bulunamadi",
               If(imgMissing = 0, "Tum " & imgTotal.ToString() & " gorselde alt etiketi var",
                  imgMissing.ToString() & "/" & imgTotal.ToString() & " gorsel alt etiketi eksik")),
            If(imgMissing = 0, "OK", "!"), If(imgMissing = 0, ColGreen, ColOrange)))

        ' 12. Robots
        Dim robotsNode = doc.DocumentNode.SelectSingleNode("//meta[@name='robots']")
        Dim robotsContent As String = If(robotsNode IsNot Nothing, robotsNode.GetAttributeValue("content", "").ToLower(), "")
        Dim isBlocked As Boolean = robotsContent.Contains("noindex")
        _checkList.Add(NewCheck("Robots Direktifi",
            If(isBlocked, "DIKKAT: noindex aktif - sayfa indekslenmez!",
               If(robotsNode IsNot Nothing, "Robots: " & robotsContent, "Varsayilan (index, follow)")),
            If(isBlocked, "!", "OK"), If(isBlocked, ColRed, ColGreen)))

        ' 13. Lang
        Dim htmlNode2 = doc.DocumentNode.SelectSingleNode("//html")
        Dim lang2 As String = If(htmlNode2 IsNot Nothing, htmlNode2.GetAttributeValue("lang", ""), "")
        _checkList.Add(NewCheck("Dil Tanimi (lang)",
            If(String.IsNullOrEmpty(lang2), "HTML lang attribute eksik", "Dil: " & lang2),
            If(String.IsNullOrEmpty(lang2), "!", "OK"),
            If(String.IsNullOrEmpty(lang2), ColOrange, ColGreen)))

        ' 14. Twitter Card
        Dim hasTw As Boolean = (doc.DocumentNode.SelectSingleNode("//meta[@name='twitter:card']") IsNot Nothing)
        _checkList.Add(NewCheck("Twitter Card",
            If(hasTw, "Twitter Card etiketi mevcut", "Twitter Card eksik"),
            If(hasTw, "OK", "!"), If(hasTw, ColGreen, ColOrange)))

        ' 15. Favicon
        Dim hasFavicon As Boolean = (doc.DocumentNode.SelectSingleNode("//link[@rel='icon' or @rel='shortcut icon']") IsNot Nothing)
        _checkList.Add(NewCheck("Favicon",
            If(hasFavicon, "Favicon tanimli", "Favicon bulunamadi"),
            If(hasFavicon, "OK", "!"), If(hasFavicon, ColGreen, ColOrange)))
    End Sub

    Private Function NewCheck(title As String, detail As String, badge As String, color As SolidColorBrush) As ChecklistItem
        Dim dim2 As SolidColorBrush = If(badge = "OK", ColGreenDim, If(badge = "X", ColRedDim, ColOrangeDim))
        Return New ChecklistItem With {
            .Title = title, .Detail = detail, .Icon = badge,
            .Badge = badge, .StatusColor = dim2, .BadgeColor = dim2, .BadgeText = color
        }
    End Function

    ' ============================================================
    '  SKOR
    ' ============================================================
    ' ============================================================
    '  SKOR (GÜNCELLENDİ: Core Web Vitals ve Mobile-First Ağırlıkları)
    ' ============================================================
    Private Sub CalculateScore(doc As HtmlDocument, url As String, ms As Long)
        Dim score As Integer = 0

        ' 1. HTTPS (+15 Puan)
        If url.StartsWith("https://") Then score += 15

        ' 2. Mobile/Viewport (+12 Puan)
        If doc.DocumentNode.SelectSingleNode("//meta[@name='viewport']") IsNot Nothing Then score += 12

        ' 3. Title Uzunluğu (+12 Puan)
        Dim titleNode = doc.DocumentNode.SelectSingleNode("//title")
        Dim titleLen As Integer = If(titleNode IsNot Nothing, titleNode.InnerText.Trim().Length, 0)
        If titleLen >= 50 AndAlso titleLen <= 60 Then
            score += 12
        ElseIf titleLen > 0 Then
            score += 6
        End If

        ' 4. Sayfa Hızı / Core Web Vitals (+10 Puan)
        If ms < 800 Then
            score += 10
        ElseIf ms < 2000 Then
            score += 5
        End If

        ' 5. H1 Varlığı (+8 Puan)
        Dim h1Nodes = doc.DocumentNode.SelectNodes("//h1")
        Dim h1Count As Integer = If(h1Nodes IsNot Nothing, h1Nodes.Count, 0)
        If h1Count = 1 Then
            score += 8
        ElseIf h1Count > 1 Then
            score += 4
        End If

        ' 6. Schema.org (+8 Puan)
        If doc.DocumentNode.SelectSingleNode("//script[@type='application/ld+json']") IsNot Nothing Then score += 8

        ' 7. Canonical (+6 Puan)
        If doc.DocumentNode.SelectSingleNode("//link[@rel='canonical']") IsNot Nothing Then score += 6


        ' === Kalan 29 Puanı Diğer SEO Dinamiklerine Dağıtıyoruz ===

        ' Meta Desc (Max 8 Puan)
        Dim descNode = doc.DocumentNode.SelectSingleNode("//meta[@name='description']")
        Dim descLen As Integer = If(descNode IsNot Nothing, descNode.GetAttributeValue("content", "").Length, 0)
        If descLen >= 150 AndAlso descLen <= 160 Then
            score += 8
        ElseIf descLen > 0 Then
            score += 4
        End If

        ' Görseller ve Alt Etiketleri (Max 7 Puan)
        Dim imgs = doc.DocumentNode.SelectNodes("//img")
        If imgs IsNot Nothing AndAlso imgs.Count > 0 Then
            Dim missing As Integer = 0
            For Each imgN As HtmlNode In imgs
                If String.IsNullOrWhiteSpace(imgN.GetAttributeValue("alt", "")) Then missing += 1
            Next
            If missing = 0 Then
                score += 7
            ElseIf missing < imgs.Count \ 2 Then
                score += 3
            End If
        Else
            score += 7 ' Görsel yoksa ceza kesilmez
        End If

        ' H2 & H3 Alt Basliklar (Max 5 Puan)
        Dim h2n = doc.DocumentNode.SelectNodes("//h2")
        Dim h3n = doc.DocumentNode.SelectNodes("//h3")
        If (If(h2n IsNot Nothing, h2n.Count, 0) + If(h3n IsNot Nothing, h3n.Count, 0)) > 0 Then score += 5

        ' Link Dengesi (Max 5 Puan)
        Dim linkNodes = doc.DocumentNode.SelectNodes("//a[@href]")
        Dim hasInternal As Boolean = False
        Dim hasExternal As Boolean = False
        If linkNodes IsNot Nothing Then
            Dim bUri As Uri = Nothing
            Uri.TryCreate(url, UriKind.Absolute, bUri)
            For Each n As HtmlNode In linkNodes
                Dim href As String = n.GetAttributeValue("href", "")
                If String.IsNullOrEmpty(href) OrElse href.StartsWith("#") OrElse href.StartsWith("javascript") Then Continue For
                If bUri IsNot Nothing Then
                    Dim abs As Uri = Nothing
                    If Uri.TryCreate(bUri, href, abs) Then
                        If abs.Host = bUri.Host Then hasInternal = True Else hasExternal = True
                    End If
                End If
            Next
        End If
        If hasInternal AndAlso hasExternal Then
            score += 5
        ElseIf hasInternal OrElse hasExternal Then
            score += 2
        End If

        ' OG Tags (Max 4 Puan)
        If doc.DocumentNode.SelectSingleNode("//meta[@property='og:title']") IsNot Nothing Then score += 4

        ' Hata payına karşı skoru 100 ile sınırla
        If score > 100 Then score = 100

        ' Arayüz Güncellemesi
        TxtScore.Text = score.ToString()
        If score >= 80 Then
            TxtScore.Foreground = ColGreen
            ScoreBadge.Background = ColGreenDim
            ScoreBadge.BorderBrush = ColGreen
        ElseIf score >= 55 Then
            TxtScore.Foreground = ColOrange
            ScoreBadge.Background = ColOrangeDim
            ScoreBadge.BorderBrush = ColOrange
        Else
            TxtScore.Foreground = ColRed
            ScoreBadge.Background = ColRedDim
            ScoreBadge.BorderBrush = ColRed
        End If
    End Sub

    ' ============================================================
    '  HTML RAPOR DISA AKTARMA
    ' ============================================================
    Private Sub BtnExport_Click(sender As Object, e As RoutedEventArgs)
        Try
            Dim sfd As New Microsoft.Win32.SaveFileDialog()
            sfd.Filter = "HTML Raporu (*.html)|*.html"
            sfd.Title = "SEO Raporunu Kaydet"
            sfd.FileName = "SEO_Raporu_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".html"

            If sfd.ShowDialog() = True Then
                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine("<!DOCTYPE html>")
                sb.AppendLine("<html lang='tr'>")
                sb.AppendLine("<head>")
                sb.AppendLine("<meta charset='UTF-8'>")
                sb.AppendLine("<title>SEO Analiz Raporu - " & WebUtility.HtmlEncode(TxtUrl.Text) & "</title>")
                sb.AppendLine("<style>")
                sb.AppendLine("body{font-family:'Segoe UI',sans-serif;background:#0A0E1A;color:#E8EDF5;margin:0;padding:40px 20px;line-height:1.6}")
                sb.AppendLine(".container{max-width:1100px;margin:auto}")
                sb.AppendLine(".header{display:flex;justify-content:space-between;align-items:center;background:#111827;padding:30px 40px;border-radius:12px;margin-bottom:30px;border:1px solid #1E2D45}")
                sb.AppendLine(".header h1{margin:0 0 10px;font-size:28px;color:#00D4FF}.header p{margin:0;color:#5A6A85;font-size:15px}")
                sb.AppendLine(".score-badge{background:#1A2235;font-size:42px;font-weight:900;padding:25px;border-radius:50%;min-width:60px;text-align:center;border:4px solid;display:flex;justify-content:center;align-items:center}")
                sb.AppendLine(".card{background:#111827;border-radius:12px;padding:30px;margin-bottom:24px;border:1px solid #1E2D45}")
                sb.AppendLine(".card h2{color:#5A6A85;font-size:14px;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid #1E2D45;padding-bottom:12px;margin-top:0;margin-bottom:20px;font-weight:700}")
                sb.AppendLine(".grid-stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:20px}")
                sb.AppendLine(".stat-box{background:#1A2235;padding:20px;border-radius:10px;text-align:center}")
                sb.AppendLine(".stat-box .val{font-size:32px;font-weight:900;margin-bottom:5px}.stat-box .lbl{color:#5A6A85;font-size:13px;font-weight:600}")
                sb.AppendLine(".table{width:100%;border-collapse:collapse}.table th,.table td{padding:16px 12px;text-align:left;border-bottom:1px solid #1E2D45;font-size:14px}")
                sb.AppendLine(".table th{color:#5A6A85;font-size:12px;font-weight:600;text-transform:uppercase}.table tr:last-child td{border-bottom:none}")
                sb.AppendLine(".badge{padding:6px 12px;border-radius:6px;font-size:11px;font-weight:bold;display:inline-block;text-align:center;min-width:45px}")
                sb.AppendLine(".kw-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:15px}")
                sb.AppendLine(".kw-item{background:#1A2235;padding:15px;border-radius:8px}")
                sb.AppendLine(".kw-header{display:flex;justify-content:space-between;margin-bottom:10px;font-size:14px;font-weight:600}")
                sb.AppendLine(".kw-bar-bg{background:#0D1525;height:6px;border-radius:3px;overflow:hidden}")
                sb.AppendLine(".kw-bar-fg{background:#00D4FF;height:100%;border-radius:3px}")
                sb.AppendLine(".footer{text-align:center;color:#5A6A85;font-size:13px;margin-top:40px;padding-top:20px;border-top:1px solid #1E2D45}")
                sb.AppendLine("</style></head><body><div class='container'>")

                Dim scoreVal As Integer = 0
                Integer.TryParse(TxtScore.Text, scoreVal)
                Dim scoreColor As String = If(scoreVal >= 80, "#00FF88", If(scoreVal >= 55, "#FF6B35", "#FF3B5C"))

                sb.AppendLine("<div class='header'><div>")
                sb.AppendLine("<h1>Kapsamli SEO Analiz Raporu</h1>")
                sb.AppendLine("<p><strong style='color:#E8EDF5;'>Hedef URL:</strong> <a href='" & TxtUrl.Text & "' target='_blank' style='color:#00D4FF;text-decoration:none;'>" & TxtUrl.Text & "</a></p>")
                sb.AppendLine("<p style='margin-top:8px;'><strong style='color:#E8EDF5;'>Olusturulma:</strong> " & DateTime.Now.ToString("dd MMMM yyyy, HH:mm") & "</p>")
                sb.AppendLine("</div><div class='score-badge' style='color:" & scoreColor & ";border-color:" & scoreColor & ";'>" & TxtScore.Text & "</div></div>")

                ' Istatistikler
                sb.AppendLine("<div class='card'><h2>Hizli Istatistikler</h2><div class='grid-stats'>")
                sb.AppendLine("<div class='stat-box'><div class='val' style='color:#00FF88;'>" & TxtSpeed.Text & "</div><div class='lbl'>Yanit Suresi</div></div>")
                sb.AppendLine("<div class='stat-box'><div class='val' style='color:#00D4FF;'>" & StatH1.Text & "</div><div class='lbl'>H1 Tag</div></div>")
                sb.AppendLine("<div class='stat-box'><div class='val' style='color:#00FF88;'>" & StatH2.Text & "</div><div class='lbl'>H2 Tag</div></div>")
                sb.AppendLine("<div class='stat-box'><div class='val' style='color:#FF6B35;'>" & StatImages.Text & "</div><div class='lbl'>Gorsel</div></div>")
                sb.AppendLine("<div class='stat-box'><div class='val' style='color:#A855F7;'>" & StatLinks.Text & "</div><div class='lbl'>Link</div></div>")
                sb.AppendLine("<div class='stat-box'><div class='val' style='color:#FFD700;'>" & StatWords.Text & "</div><div class='lbl'>Kelime</div></div>")
                sb.AppendLine("</div></div>")

                ' Kontrol listesi
                If _checkList.Count > 0 Then
                    sb.AppendLine("<div class='card'><h2>SEO Kontrol Listesi</h2>")
                    sb.AppendLine("<table class='table'><tr><th style='width:80px;'>Durum</th><th>Kriter</th><th>Sonuc</th></tr>")
                    For Each item As ChecklistItem In _checkList
                        Dim bs As String = If(item.Badge = "OK", "background:rgba(0,255,136,.15);color:#00FF88;",
                                            If(item.Badge = "X", "background:rgba(255,59,92,.15);color:#FF3B5C;",
                                            "background:rgba(255,107,53,.15);color:#FF6B35;"))
                        sb.AppendLine("<tr><td><span class='badge' style='" & bs & "'>" & item.Badge & "</span></td>")
                        sb.AppendLine("<td style='font-weight:600;color:#E8EDF5;'>" & WebUtility.HtmlEncode(item.Title) & "</td>")
                        sb.AppendLine("<td style='color:#5A6A85;'>" & WebUtility.HtmlEncode(item.Detail) & "</td></tr>")
                    Next
                    sb.AppendLine("</table></div>")
                End If

                ' Anahtar kelimeler
                If _keywordList.Count > 0 Then
                    sb.AppendLine("<div class='card'><h2>Anahtar Kelime Yogunlugu</h2><div class='kw-grid'>")
                    For Each kw As KeywordItem In _keywordList
                        Dim pct As String = kw.Percent.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture)
                        sb.AppendLine("<div class='kw-item'><div class='kw-header'><span style='color:#E8EDF5;'>" & WebUtility.HtmlEncode(kw.Word) & "</span><span style='color:#5A6A85;'>" & kw.CountText & "</span></div>")
                        sb.AppendLine("<div class='kw-bar-bg'><div class='kw-bar-fg' style='width:" & pct & "%;'></div></div></div>")
                    Next
                    sb.AppendLine("</div></div>")
                End If

                ' Teknik SEO
                If _techList.Count > 0 Then
                    sb.AppendLine("<div class='card'><h2>Teknik SEO Verileri</h2>")
                    sb.AppendLine("<table class='table'><tr><th style='width:200px;'>Kriter</th><th>Deger</th><th style='width:100px;text-align:right;'>Durum</th></tr>")
                    For Each tech As TechItem In _techList
                        Dim isGood As Boolean = (tech.Status = "OK" OrElse tech.Status = "TAMAM" OrElse tech.Status = "VAR" OrElse tech.Status = "HTTPS" OrElse tech.Status = "IYI")
                        Dim isBad As Boolean = (tech.Status = "EKSIK" OrElse tech.Status = "X" OrElse tech.Status = "HTTP" OrElse tech.Status = "NOINDEX")
                        Dim bs As String = If(isGood, "background:rgba(0,255,136,.15);color:#00FF88;",
                                            If(isBad, "background:rgba(255,59,92,.15);color:#FF3B5C;",
                                            "background:rgba(255,107,53,.15);color:#FF6B35;"))
                        sb.AppendLine("<tr><td style='color:#5A6A85;font-weight:600;'>" & WebUtility.HtmlEncode(tech.Label) & "</td>")
                        sb.AppendLine("<td style='color:#E8EDF5;'>" & WebUtility.HtmlEncode(tech.Value) & "</td>")
                        sb.AppendLine("<td style='text-align:right;'><span class='badge' style='" & bs & "'>" & tech.Status & "</span></td></tr>")
                    Next
                    sb.AppendLine("</table></div>")
                End If

                sb.AppendLine("<div class='footer'>Bu rapor <b>SEO Analyzer Pro</b> tarafindan olusturulmustur. &copy; " & DateTime.Now.Year & "</div>")
                sb.AppendLine("</div></body></html>")

                System.IO.File.WriteAllText(sfd.FileName, sb.ToString(), System.Text.Encoding.UTF8)
                TxtStatus.Text = "HTML Rapor kaydedildi: " & sfd.FileName

                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                    .FileName = sfd.FileName,
                    .UseShellExecute = True
                })
            End If
        Catch ex As Exception
            MessageBox.Show("Rapor kaydedilirken hata: " & ex.Message, "Hata", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' ============================================================
    '  YARDIMCILAR
    ' ============================================================
    Private Sub SetBadge(badge As System.Windows.Controls.Border,
                         lbl As System.Windows.Controls.TextBlock,
                         text As String, bg As SolidColorBrush, fg As SolidColorBrush)
        badge.Background = bg
        lbl.Text = text
        lbl.Foreground = fg
    End Sub

    Private Function GetCleanText(node As HtmlNode) As String
        If node Is Nothing Then Return ""
        Dim sb As New System.Text.StringBuilder()
        Dim textNodes = node.SelectNodes("//text()[not(ancestor::script) and not(ancestor::style) and not(ancestor::noscript) and not(ancestor::svg)]")
        If textNodes IsNot Nothing Then
            For Each n As HtmlNode In textNodes
                Dim txt = WebUtility.HtmlDecode(n.InnerText).Trim()
                If Not String.IsNullOrWhiteSpace(txt) Then
                    sb.AppendLine(txt)
                End If
            Next
        End If
        Return sb.ToString()
    End Function

    Private Sub ResetUI()
        TxtTitle.Text = "Analiz bekleniyor..."
        TxtDesc.Text = "Analiz bekleniyor..."
        TxtTitleLen.Text = "0 karakter"
        TxtDescLen.Text = "0 karakter"
        TitleProgress.Value = 0
        DescProgress.Value = 0
        StatH1.Text = "0"
        StatH2.Text = "0"
        StatImages.Text = "0"
        StatLinks.Text = "0"
        StatWords.Text = "0"
        TxtSpeed.Text = "-- ms"
        TxtSpeed.Foreground = ColTextMuted
        TxtSpeedNote.Text = "Yanit suresi"
        TxtScore.Text = "--"
        TxtScore.Foreground = ColTextMuted
        ScoreBadge.Background = ColCardBg
        ScoreBadge.BorderBrush = New SolidColorBrush(Color.FromRgb(30, 45, 69))
        TxtInternalCount.Text = "0"
        TxtExternalCount.Text = "0"
        TxtAltOkCount.Text = "0"
        TxtAltMissingCount.Text = "0"
        TxtBrokenCount.Text = "0"
        BtnExport.IsEnabled = False

        ' ObservableCollection'lari temizle
        _headingList.Clear()
        _linkList.Clear()
        _imageList.Clear()
        _checkList.Clear()
        _techList.Clear()
        _keywordList.Clear()
        _robotsList.Clear()
        _compareList.Clear()

        ' Icerik analizi UI sifirla
        TxtReadScore.Text = "--"
        TxtReadLevel.Text = "Analiz bekleniyor"
        TxtReadDetail.Text = ""
        ReadProgress.Value = 0
        TxtTextRatioPct.Text = "0%"
        TxtHtmlSize.Text = "0 KB"
        TxtTextSize.Text = "0 KB"
        TextRatioBar.Width = 0
        StatReadability.Text = "--"
        StatTextRatio.Text = "--"
        TxtReadabilityNote.Text = "--"
        BtnExportCsv.IsEnabled = False

        SetBadge(TitleStatusBadge, TitleStatus, "--", ColCardBg, ColTextMuted)
        SetBadge(DescStatusBadge, DescStatus, "--", ColCardBg, ColTextMuted)

        Application.Current.Dispatcher.InvokeAsync(Sub()
                                                       TxtRawHtml.Clear()
                                                   End Sub, Threading.DispatcherPriority.Background)
    End Sub
    ' ============================================================
    '  OKUNABİLİRLİK ANALİZİ (Türkçe - Ateşman Formülü)
    ' ============================================================
    Private Sub AnalyzeReadability(doc As HtmlDocument)
        Dim bodyNode = doc.DocumentNode.SelectSingleNode("//body")
        If bodyNode Is Nothing Then Return

        Dim cleanText As String = GetCleanText(bodyNode)

        ' Sadece harfler, sayılar, boşluklar (Enter dahil) ve noktalama işaretleri kalsın
        Dim rgx As New System.Text.RegularExpressions.Regex("[^\p{L}\d\s\.\!\?]")
        cleanText = rgx.Replace(cleanText, " ")

        ' KRİTİK DÜZELTME 1: Noktalama işaretlerine ek olarak HTML satır sonlarını (\n) da cümle bitişi sayıyoruz.
        ' Yoksa alt alta yazılmış menü linklerini 500 kelimelik tek bir cümle sanır!
        Dim sentenceRgx As New System.Text.RegularExpressions.Regex("[\.\!\?\n]+")
        Dim sentences As Integer = sentenceRgx.Matches(cleanText).Count
        If sentences = 0 Then sentences = 1

        ' Kelimeleri ayır
        Dim words() As String = cleanText.Split(New Char() {" "c, Chr(13), Chr(10), Chr(9)}, StringSplitOptions.RemoveEmptyEntries)
        Dim wordCount As Integer = 0
        For Each w As String In words
            If w.Trim().Length > 1 Then wordCount += 1
        Next
        If wordCount = 0 Then Return

        ' Ortalama kelime/cümle (Ateşman Formülü: x2)
        Dim avgWordsPerSentence As Double = CDbl(wordCount) / CDbl(sentences)

        ' Hece hesaplama (Türkçe sesli harfler: a, e, ı, i, o, ö, u, ü)
        Dim vowels As Integer = 0
        Dim vowelRgx As New System.Text.RegularExpressions.Regex("[aeıioöuüAEIİOÖUÜ]")
        For Each w As String In words
            Dim syllables As Integer = vowelRgx.Matches(w).Count
            If syllables = 0 Then syllables = 1
            vowels += syllables
        Next

        ' Ortalama hece/kelime (Ateşman Formülü: x1)
        Dim avgSyllables As Double = CDbl(vowels) / CDbl(wordCount)

        ' KRİTİK DÜZELTME 2: Türkçe için Ateşman Okunabilirlik Formülü
        Dim score As Double = 198.825 - (40.175 * avgSyllables) - (2.61 * avgWordsPerSentence)

        If score > 100 Then score = 100
        If score < 0 Then score = 0
        Dim scoreInt As Integer = CInt(Math.Round(score))

        Dim level As String
        Dim detail As String
        Dim fg As SolidColorBrush

        ' Ateşman Değerlendirme Ölçeği
        If scoreInt >= 70 Then
            level = "Kolay Okunur"
            detail = "Kısa ve net kelimeler. İlkokul/Ortaokul düzeyi için uygundur."
            fg = ColGreen
        ElseIf scoreInt >= 50 Then
            level = "Orta Düzey"
            detail = "Standart Türkçe metni. Lise düzeyi okuyucu için idealdir."
            fg = ColCyanDim ' Mavi tonu arayüzüne daha şık gider
        ElseIf scoreInt >= 30 Then
            level = "Zor Okunur"
            detail = "Uzun kelimeler ve karmaşık cümleler. Lisans/Akademik düzey."
            fg = ColOrange
        Else
            level = "Çok Zor"
            detail = "Ağır bir dili var. Okumayı ve taramayı zorlaştırıyor."
            fg = ColRed
        End If

        ' ARAYÜZ (UI) GÜNCELLEMESİ
        Application.Current.Dispatcher.Invoke(Sub()
                                                  If TxtReadScore IsNot Nothing Then
                                                      TxtReadScore.Text = scoreInt.ToString()
                                                      TxtReadScore.Foreground = fg
                                                      TxtReadLevel.Text = level
                                                      TxtReadLevel.Foreground = fg
                                                      TxtReadDetail.Text = detail

                                                      If ReadProgress IsNot Nothing Then
                                                          ReadProgress.Value = scoreInt
                                                          ReadProgress.Foreground = fg
                                                      End If

                                                      If StatReadability IsNot Nothing Then
                                                          StatReadability.Text = scoreInt.ToString()
                                                          StatReadability.Foreground = fg
                                                      End If

                                                      If TxtReadabilityNote IsNot Nothing Then
                                                          TxtReadabilityNote.Text = level
                                                      End If
                                                  End If
                                              End Sub)
    End Sub

    ' ============================================================
    '  METIN / HTML ORANI
    ' ============================================================
    Private Sub AnalyzeTextHtmlRatio(doc As HtmlDocument, rawHtml As String)
        Dim bodyNode = doc.DocumentNode.SelectSingleNode("//body")
        If bodyNode Is Nothing Then Return

        Dim htmlBytes As Integer = System.Text.Encoding.UTF8.GetByteCount(rawHtml)
        Dim textContent As String = GetCleanText(bodyNode)
        Dim textBytes As Integer = System.Text.Encoding.UTF8.GetByteCount(textContent)

        Dim htmlKb As Double = htmlBytes / 1024.0
        Dim textKb As Double = textBytes / 1024.0
        Dim ratio As Double = If(htmlBytes > 0, (CDbl(textBytes) / CDbl(htmlBytes)) * 100.0, 0)

        TxtHtmlSize.Text = CInt(htmlKb).ToString() & " KB"
        TxtTextSize.Text = CInt(textKb).ToString() & " KB"
        TxtTextRatioPct.Text = CInt(ratio).ToString() & "%"
        StatTextRatio.Text = CInt(ratio).ToString() & "%"

        ' Bar genisligi hesapla (max 600px gibi gorsel genisligin %80'i)
        Dim maxBarWidth As Double = 500
        Dim barWidth As Double = Math.Min(ratio / 100.0 * maxBarWidth, maxBarWidth)
        TextRatioBar.Width = barWidth

        If ratio >= 25 Then
            TxtTextRatioPct.Foreground = ColGreen
            TextRatioBar.Background = ColGreen
        ElseIf ratio >= 10 Then
            TxtTextRatioPct.Foreground = ColYellow
            TextRatioBar.Background = ColYellow
        Else
            TxtTextRatioPct.Foreground = ColRed
            TextRatioBar.Background = ColRed
        End If
    End Sub

    ' ============================================================
    '  ROBOTS.TXT & SITEMAP KONTROLU
    ' ============================================================
    Private Async Function AnalyzeRobotsAndSitemapAsync(url As String) As Task
        Dim baseUri As Uri = Nothing
        If Not Uri.TryCreate(url, UriKind.Absolute, baseUri) Then Return

        Dim base As String = baseUri.Scheme & "://" & baseUri.Host

        Using client As New HttpClient()
            client.Timeout = TimeSpan.FromSeconds(8)
            client.DefaultRequestHeaders.Add("User-Agent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

            ' robots.txt
            Dim robotsOk As Boolean = False
            Dim robotsContent As String = "(okunamadi)"
            Try
                Dim resp = Await client.GetAsync(base & "/robots.txt")
                If resp.IsSuccessStatusCode Then
                    robotsOk = True
                    Dim txt As String = Await resp.Content.ReadAsStringAsync()
                    Dim lines() As String = txt.Split(New Char() {Chr(10)}, StringSplitOptions.RemoveEmptyEntries)
                    robotsContent = "OK - " & lines.Length.ToString() & " satir"
                Else
                    robotsContent = "HTTP " & CInt(resp.StatusCode).ToString()
                End If
            Catch
                robotsContent = "Erisim hatasi"
            End Try
            _robotsList.Add(MakeTechItem("robots.txt", base & "/robots.txt  |  " & robotsContent,
                If(robotsOk, "OK", "EKSIK"),
                If(robotsOk, ColGreenDim, ColOrangeDim),
                If(robotsOk, ColGreen, ColOrange)))

            ' sitemap.xml
            Dim sitemapOk As Boolean = False
            Dim sitemapContent As String = "(okunamadi)"
            Try
                Dim resp2 = Await client.GetAsync(base & "/sitemap.xml")
                If resp2.IsSuccessStatusCode Then
                    sitemapOk = True
                    Dim txt2 As String = Await resp2.Content.ReadAsStringAsync()
                    Dim urlCount As Integer = System.Text.RegularExpressions.Regex.Matches(txt2, "<url>").Count
                    sitemapContent = "OK" & If(urlCount > 0, " - " & urlCount.ToString() & " URL", "")
                Else
                    sitemapContent = "HTTP " & CInt(resp2.StatusCode).ToString()
                End If
            Catch
                sitemapContent = "Erisim hatasi"
            End Try
            _robotsList.Add(MakeTechItem("sitemap.xml", base & "/sitemap.xml  |  " & sitemapContent,
                If(sitemapOk, "OK", "EKSIK"),
                If(sitemapOk, ColGreenDim, ColOrangeDim),
                If(sitemapOk, ColGreen, ColOrange)))

            ' Gzip / sikistirma
            Dim gzipOk As Boolean = False
            Try
                Dim req As New HttpRequestMessage(HttpMethod.Get, url)
                req.Headers.Add("Accept-Encoding", "gzip, deflate")
                Dim resp3 = Await client.SendAsync(req)
                Dim enc As String = ""
                If resp3.Content.Headers.ContentEncoding IsNot Nothing Then
                    For Each e As String In resp3.Content.Headers.ContentEncoding
                        enc &= e & " "
                    Next
                End If
                gzipOk = enc.ToLower().Contains("gzip") OrElse enc.ToLower().Contains("br")
                _robotsList.Add(MakeTechItem("Gzip / Sikistirma",
                    If(gzipOk, "Aktif (" & enc.Trim() & ")", "Aktif degil - sayfa boyutunu azaltir"),
                    If(gzipOk, "AKTIF", "KAPALI"),
                    If(gzipOk, ColGreenDim, ColOrangeDim),
                    If(gzipOk, ColGreen, ColOrange)))
            Catch
                _robotsList.Add(MakeTechItem("Gzip / Sikistirma", "Kontrol edilemedi",
                    "?", ColOrangeDim, ColOrange))
            End Try
        End Using
    End Function

    ' ============================================================
    '  RAKIP KARSILASTIRMA
    ' ============================================================
    Private Async Function AnalyzeRakipAsync(rakipUrl As String) As Task
        Dim handler As New HttpClientHandler()
        handler.AllowAutoRedirect = True
        Using client As New HttpClient(handler)
            client.Timeout = TimeSpan.FromSeconds(20)
            client.DefaultRequestHeaders.Add("User-Agent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

            Dim sw As New System.Diagnostics.Stopwatch()
            sw.Start()
            Dim html As String = Await client.GetStringAsync(rakipUrl)
            sw.Stop()

            Dim doc2 As New HtmlDocument()
            doc2.LoadHtml(html)

            ' Karsilastirma fonksiyonu
            Dim AddRow As Action(Of String, String, String) =
                Sub(label As String, v1 As String, v2 As String)
                    ' Kim kazaniyor? (daha uzun/daha cok icerik genellikle iyi)
                    Dim winner As String = "ESIT"
                    Dim winColor As SolidColorBrush = ColCyanDim
                    Dim winFg As SolidColorBrush = ColCyan
                    Dim n1 As Integer = 0
                    Dim n2 As Integer = 0
                    If Integer.TryParse(System.Text.RegularExpressions.Regex.Match(v1, "\d+").Value, n1) AndAlso
                       Integer.TryParse(System.Text.RegularExpressions.Regex.Match(v2, "\d+").Value, n2) Then
                        If n1 > n2 Then
                            winner = "SENDE"
                            winColor = ColGreenDim
                            winFg = ColGreen
                        ElseIf n2 > n1 Then
                            winner = "RAKIP"
                            winColor = ColRedDim
                            winFg = ColRed
                        End If
                    End If
                    _compareList.Add(New CompareItem With {
                        .Label = label, .Val1 = v1, .Val2 = v2,
                        .Winner = winner, .WinColor = winColor, .WinFg = winFg
                    })
                End Sub

            ' Title
            Dim titleNode1 = _currentDoc.DocumentNode.SelectSingleNode("//title")
            Dim titleNode2 = doc2.DocumentNode.SelectSingleNode("//title")
            Dim titleTxt1 As String = If(titleNode1 IsNot Nothing, titleNode1.InnerText.Trim(), "(yok)")
            Dim titleTxt2 As String = If(titleNode2 IsNot Nothing, titleNode2.InnerText.Trim(), "(yok)")
            AddRow("Baslik", titleTxt1.Length.ToString() & " kar | " & If(titleTxt1.Length > 50, titleTxt1.Substring(0, 50) & "...", titleTxt1),
                             titleTxt2.Length.ToString() & " kar | " & If(titleTxt2.Length > 50, titleTxt2.Substring(0, 50) & "...", titleTxt2))

            ' Meta Desc
            Dim d1 = _currentDoc.DocumentNode.SelectSingleNode("//meta[@name='description']")
            Dim d2 = doc2.DocumentNode.SelectSingleNode("//meta[@name='description']")
            Dim dv1 As Integer = If(d1 IsNot Nothing, d1.GetAttributeValue("content", "").Length, 0)
            Dim dv2 As Integer = If(d2 IsNot Nothing, d2.GetAttributeValue("content", "").Length, 0)
            AddRow("Meta Aciklama", dv1.ToString() & " karakter", dv2.ToString() & " karakter")

            ' H1 sayisi
            Dim h1a = _currentDoc.DocumentNode.SelectNodes("//h1")
            Dim h1b = doc2.DocumentNode.SelectNodes("//h1")
            AddRow("H1 Sayisi",
                If(h1a IsNot Nothing, h1a.Count.ToString(), "0"),
                If(h1b IsNot Nothing, h1b.Count.ToString(), "0"))

            ' H2 sayisi
            Dim h2a = _currentDoc.DocumentNode.SelectNodes("//h2")
            Dim h2b = doc2.DocumentNode.SelectNodes("//h2")
            AddRow("H2 Sayisi",
                If(h2a IsNot Nothing, h2a.Count.ToString(), "0"),
                If(h2b IsNot Nothing, h2b.Count.ToString(), "0"))

            ' Link sayisi
            Dim la = _currentDoc.DocumentNode.SelectNodes("//a[@href]")
            Dim lb = doc2.DocumentNode.SelectNodes("//a[@href]")
            AddRow("Toplam Link",
                If(la IsNot Nothing, la.Count.ToString(), "0"),
                If(lb IsNot Nothing, lb.Count.ToString(), "0"))

            ' Gorsel sayisi
            Dim ia = _currentDoc.DocumentNode.SelectNodes("//img")
            Dim ib = doc2.DocumentNode.SelectNodes("//img")
            AddRow("Gorsel Sayisi",
                If(ia IsNot Nothing, ia.Count.ToString(), "0"),
                If(ib IsNot Nothing, ib.Count.ToString(), "0"))

            ' Kelime sayisi (yaklasik)
            Dim body1 = _currentDoc.DocumentNode.SelectSingleNode("//body")
            Dim body2 = doc2.DocumentNode.SelectSingleNode("//body")
            Dim wc1 As Integer = 0
            Dim wc2 As Integer = 0
            If body1 IsNot Nothing Then
                Dim rgx As New System.Text.RegularExpressions.Regex("[^\p{L}\s]")
                Dim t As String = rgx.Replace(GetCleanText(body1), " ")
                wc1 = t.Split(New Char() {" "c, Chr(13), Chr(10)}, StringSplitOptions.RemoveEmptyEntries).Length
            End If
            If body2 IsNot Nothing Then
                Dim rgx2 As New System.Text.RegularExpressions.Regex("[^\p{L}\s]")
                Dim bodyText2 As String = rgx2.Replace(GetCleanText(body2), " ")
                wc2 = bodyText2.Split(New Char() {" "c, Chr(13), Chr(10)}, StringSplitOptions.RemoveEmptyEntries).Length
            End If
            AddRow("Kelime Sayisi", wc1.ToString(), wc2.ToString())

            ' Schema.org
            Dim s1 = _currentDoc.DocumentNode.SelectNodes("//script[@type='application/ld+json']")
            Dim s2 = doc2.DocumentNode.SelectNodes("//script[@type='application/ld+json']")
            AddRow("Schema.org Blogu",
                If(s1 IsNot Nothing, s1.Count.ToString(), "0"),
                If(s2 IsNot Nothing, s2.Count.ToString(), "0"))

            ' OG Tags
            Dim og1 As Integer = 0
            Dim og2 As Integer = 0
            For Each prop As String In New String() {"og:title", "og:description", "og:image"}
                If _currentDoc.DocumentNode.SelectSingleNode("//meta[@property='" & prop & "']") IsNot Nothing Then og1 += 1
                If doc2.DocumentNode.SelectSingleNode("//meta[@property='" & prop & "']") IsNot Nothing Then og2 += 1
            Next
            AddRow("Open Graph Etiketi", og1.ToString() & "/3", og2.ToString() & "/3")

            ' HTML boyutu
            AddRow("HTML Boyutu",
                CInt(_currentDoc.DocumentNode.OuterHtml.Length / 1024).ToString() & " KB",
                CInt(doc2.DocumentNode.OuterHtml.Length / 1024).ToString() & " KB")

            ' Sayfa hizi
            AddRow("Yanit Suresi (ms)", _currentSpeedMs, sw.ElapsedMilliseconds.ToString() & " ms")

            ' Genel SEO skoru (mevcut skor vs rakip hesapla)
            Dim rakipScore As Integer = 0
            Dim rakipTitleLen As Integer = If(titleNode2 IsNot Nothing, titleNode2.InnerText.Trim().Length, 0)
            If rakipTitleLen >= 50 Then
                rakipScore += 15
            ElseIf rakipTitleLen > 0 Then
                rakipScore += 7
            End If
            If dv2 >= 150 Then rakipScore += 10 Else If dv2 > 0 Then rakipScore += 5
            If (If(h1b IsNot Nothing, h1b.Count, 0)) = 1 Then rakipScore += 10
            If doc2.DocumentNode.SelectSingleNode("//meta[@name='viewport']") IsNot Nothing Then rakipScore += 5
            If rakipUrl.StartsWith("https://") Then rakipScore += 10
            If doc2.DocumentNode.SelectSingleNode("//link[@rel='canonical']") IsNot Nothing Then rakipScore += 5
            Dim myScore As Integer = 0
            Integer.TryParse(TxtScore.Text, myScore)
            AddRow("SEO Skoru (tahmini)", myScore.ToString() & " puan", rakipScore.ToString() & " puan")
        End Using
    End Function

    ' ============================================================
    '  GECMIS KAYDI
    ' ============================================================
    Private Sub AddToHistory(url As String, ms As Long)
        Dim scoreVal As Integer = 0
        Integer.TryParse(TxtScore.Text, scoreVal)
        Dim scoreColor As SolidColorBrush = If(scoreVal >= 80, ColGreenDim, If(scoreVal >= 55, ColOrangeDim, ColRedDim))
        Dim scoreFg As SolidColorBrush = If(scoreVal >= 80, ColGreen, If(scoreVal >= 55, ColOrange, ColRed))

        Dim displayUrl As String = If(url.Length > 60, url.Substring(0, 60) & "...", url)

        ' Ayni URL varsa guncelle, yoksa basa ekle
        Dim existing As HistoryItem = Nothing
        For Each item As HistoryItem In _historyList
            If item.Url = displayUrl Then
                existing = item
                Exit For
            End If
        Next
        If existing IsNot Nothing Then
            _historyList.Remove(existing)
        End If

        _historyList.Insert(0, New HistoryItem With {
            .Url = displayUrl,
            .AnalysisDate = DateTime.Now.ToString("dd.MM.yyyy HH:mm"),
            .Score = TxtScore.Text,
            .Speed = ms.ToString() & " ms",
            .ScoreColor = scoreColor,
            .ScoreFg = scoreFg
        })

        ' Gecmisi 20 kayitla sinirla
        Do While _historyList.Count > 20
            _historyList.RemoveAt(_historyList.Count - 1)
        Loop
        SaveHistory()
    End Sub
    ' Geçmiş dosyasının bilgisayarda kaydedileceği yol (Uygulamanın çalıştığı klasör)
    Private ReadOnly HistoryFilePath As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "seo_history.txt")

    ' ============================================================
    '  GEÇMİŞİ DOSYAYA KAYDET
    ' ============================================================
    Private Sub SaveHistory()
        Try
            Dim sb As New System.Text.StringBuilder()
            For Each item As HistoryItem In _historyList
                ' Verileri aralarına "|" koyarak tek satırda birleştiriyoruz
                sb.AppendLine($"{item.Url}|{item.AnalysisDate}|{item.Score}|{item.Speed}")
            Next
            System.IO.File.WriteAllText(HistoryFilePath, sb.ToString(), System.Text.Encoding.UTF8)
        Catch ex As Exception
            ' Kaydetme hatası olursa program çökmesin diye boş bırakıyoruz
        End Try
    End Sub

    ' ============================================================
    '  GEÇMİŞİ DOSYADAN YÜKLE (Açılışta çalışacak)
    ' ============================================================
    Private Sub LoadHistory()
        Try
            If System.IO.File.Exists(HistoryFilePath) Then
                Dim lines As String() = System.IO.File.ReadAllLines(HistoryFilePath, System.Text.Encoding.UTF8)
                _historyList.Clear()

                ' HATA BURADAYDI: "line" yerine "strLine" kullandık ve String olduğunu belirttik
                For Each strLine As String In lines
                    If String.IsNullOrWhiteSpace(strLine) Then Continue For

                    Dim parts() As String = strLine.Split("|"c)

                    If parts.Length >= 4 Then
                        Dim scoreVal As Integer = 0
                        Integer.TryParse(parts(2), scoreVal)

                        ' Renkleri skora göre yeniden hesapla
                        Dim scoreColor As SolidColorBrush = If(scoreVal >= 80, ColGreenDim, If(scoreVal >= 55, ColOrangeDim, ColRedDim))
                        Dim scoreFg As SolidColorBrush = If(scoreVal >= 80, ColGreen, If(scoreVal >= 55, ColOrange, ColRed))

                        _historyList.Add(New HistoryItem With {
                            .Url = parts(0),
                            .AnalysisDate = parts(1),
                            .Score = parts(2),
                            .Speed = parts(3),
                            .ScoreColor = scoreColor,
                            .ScoreFg = scoreFg
                        })
                    End If
                Next
            End If
        Catch ex As Exception
            ' Yükleme hatası olursa yoksay
        End Try
    End Sub
End Class
