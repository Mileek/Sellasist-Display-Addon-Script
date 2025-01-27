/** 
 * Zmienne Globalne:
 * @param {string} WEBSITE_NAME - Nazwa strony internetowej, u Was to będzie "websiteName"
 * @param {string} ALLEGRO_NOTIFICATION - Nazwa nadawcy maila z Allegro, u Was to będzie "powiadomienia@allegro.pl
 */
var WEBSITE_NAME = "websiteName";
var ALLEGRO_NOTIFICATION = "powiadomienia@allegro.pl";

/**
 * Główna funkcja add-onu wywoływana przy otwarciu wiadomości Gmail, event otwierający emaila
 * @param {Object} e - Obiekt zawierający metadane wiadomości
 * @returns {Card} Karta z informacjami o zamówieniach
 */
function onGmailMessage(e)
{
    // Pobierz dane wiadomości
    var messageId = e.messageMetadata.messageId;
    var accessToken = e.messageMetadata.accessToken;

    // Konfiguracja dostępu do Gmail API
    GmailApp.setCurrentMessageAccessToken(accessToken);
    var message = GmailApp.getMessageById(messageId);

    // Wydobycie adresu email nadawcy
    var fromHeader = message.getFrom();
    var emailRegex = /<([^>]+)>/;
    var match = fromHeader.match(emailRegex);
    var senderEmail = match ? match[1] : fromHeader;

    //TODO: Dodać normalny wybór dat przez użytkownika
    // Konfiguracja dat dla filtrowania zamówień, pobiera od roku 2024 do dzisiejszej daty i godziny, 
    var filterDates = getFilterDate();
    var dateFrom = filterDates[0];
    var dateTo = filterDates[1];

    //Sprawdźczy mail jest od allegro
    if (senderEmail === ALLEGRO_NOTIFICATION)
    {
        // Wydobycie adresu email bez "+identyfikator"
        var replyToHeader = message.getReplyTo();
        var mathReply = replyToHeader.match(emailRegex)
        var replyToEmail = mathReply ? mathReply[1] : replyToHeader;

        // Znajdź wszystkie powiązane maile Allegro
        var allegroEmails = findAllegroEmails(replyToEmail);

        // Pobierz dane dla każdego znalezionego maila
        var sellasistData = [];

        allegroEmails.forEach(function (email)
        {
            var dataForEmail = getSellasistData(email, dateFrom, dateTo);

            sellasistData = sellasistData.concat(dataForEmail.orders);
        })

        // Pobranie danych z Sellasist i utworzenie karty
        return buildSellasistCard(replyToEmail, {
            orders: sellasistData,
        });
    }
    else
    {
        // Pobranie danych z Sellasist i utworzenie karty
        var sellasistData = getSellasistData(senderEmail, dateFrom, dateTo);
        return buildSellasistCard(senderEmail, sellasistData);
    }
}

/**
 * Pobiera zakres dat do filtrowania zamówień w API Sellasist
 * Obecnie ustawiona data od 1 grudnia 2024 do dnia dzisiejszego * 
 * @returns {Array} Tablica zawierająca:
 *   - dateFrom: Data początkowa w formacie 'YYYY-MM-DD HH:mm:ss'
 *   - dateTo: Data końcowa (dzisiejsza) w formacie 'YYYY-MM-DD HH:mm:ss'
 */
function getFilterDate()
{
    var currentDate = new Date();
    var pastDate = new Date('2024-12-01');
    var dateFrom = pastDate.toISOString().replace('T', ' ').split('.')[0];
    var dateTo = currentDate.toISOString().replace('T', ' ').split('.')[0];
    return [dateFrom, dateTo];
}

/**
 * Funkcja odpowiedzialna za przeszukiwanie skrzynki odbiorczej w poszukiwaniu adresu email w formacie baseEmail+identyfikator@allegromail
 * @param {string} replyToEmail - Obiekt zawierający metadane wiadomości
 */
function findAllegroEmails(replyToEmail)
{
    // podziel email tak żeby dostać to co przed @
    var baseEmail = replyToEmail.split('@')[0];
    //Szukaj emailu po regexie, zaczynającym się od początku maila, oddzielonego @, potem może być "+ i cokolwiek" i zakończone na pl ALBO com, jak będzie potrzebna obsługa od de, albo cz wystarczy to dodać pl|com|de
    var emailPattern = new RegExp('^' + baseEmail + '\\+[a-zA-Z0-9]+@allegromail\\.(pl|com)$');

    // Pobierz wątki związane z mailem dla całego emaila (czyli będzie ich mało), przeszukaj max 20 wiadomości
    var threads = GmailApp.search(replyToEmail, 0, 20);
    var allegroEmails = [];
    var uniqueEmails = {}; // Do sprawdzania duplikatów

    for (var i = 0; i < threads.length; i++)
    {
        var messages = threads[i].getMessages();
        // Pobierna replyTo z pierwszej wiadomości w wątku, ponieważ reply-to powinien być zawsze ten sam
        var messageReplyTo = messages[0].getReplyTo();
        var matches = messageReplyTo.match(emailPattern);

        // Sprawdzaj czy email jest zgodny z wzorcem i czy nie jest duplikatem
        if (matches && !uniqueEmails[matches[0]])
        {
            uniqueEmails[matches[0]] = true;
            allegroEmails.push(matches[0]);
        }
    }

    return allegroEmails;
}

/**
 * Pobiera dane zamówień z API Sellasist dla podanego adresu email
 * @param {string} email - Adres email klienta
 * @param {string} dateFrom - Filtr Od
 * @param {string} dateTo - Filtr Do
 * @returns {Object} Obiekt zawierający zamówienia lub informację o błędzie
 */
function getSellasistData(email, dateFrom, dateTo)
{
    var sellasistApiKey = getSellasistApiKey();

    if (!sellasistApiKey)
    {
        return { error: "Brak klucza SELLASIST_API_KEY w Script Properties" };
    }

    var url = createFetchUrl(email, dateFrom, dateTo);

    // Konfiguracja zapytania HTTP
    var options = {
        method: 'get',
        headers: {
            'apiKey': sellasistApiKey,
            'accept': 'application/json'
        },
        muteHttpExceptions: true
    };

    try
    {
        // Wykonanie zapytania do API
        var response = UrlFetchApp.fetch(url, options);
        var statusCode = response.getResponseCode();
        var contentText = response.getContentText();

        if (statusCode === 200)
        {
            var json = JSON.parse(contentText);
            return { orders: Array.isArray(json) ? json : [] };
        } else
        {
            return { error: "Sellasist: błąd " + statusCode + " - " + contentText };
        }
    } catch (err)
    {
        return { error: "Błąd wywołania Sellasist: " + err.toString() };
    }
}

/**
 * Pobiera klucz API Sellasist z ustawień skryptu, https://api.sellasist.pl/ trzeba się zautoryzować żeby korzystać z poprawnych requestów
 * @returns {string} Klucz API Sellasist
 */
function getSellasistApiKey()
{
    return PropertiesService.getScriptProperties().getProperty("SELLASIST_API_KEY");
}

/**
 * Tworzy URL do API Sellasist z odpowiednimi parametrami do filtrowania zamówień
 * @param {string} email - Adres email klienta
 * @param {string} dateFrom - Data początkowa w formacie 'YYYY-MM-DD HH:mm:ss'
 * @param {string} dateTo - Data końcowa w formacie 'YYYY-MM-DD HH:mm:ss'
 * @returns {string} URL z parametrami do API Sellasist
 */
function createFetchUrl(email, dateFrom, dateTo)
{
    return "https://" + WEBSITE_NAME + ".sellasist.pl/api/v1/orders" +
        "?offset=0" +
        "&limit=50" +
        "&email=" + encodeURIComponent(email) +
        "&date_from=" + encodeURIComponent(dateFrom) +
        "&date_to=" + encodeURIComponent(dateTo);
}

/**
 * Tworzy kartę z informacjami o zamówieniach dla add-onu Gmail
 * @param {string} senderEmail - Adres email nadawcy
 * @param {Object} sellasistData - Dane zamówień z Sellasist
 * @returns {Card} Karta do wyświetlenia w Gmail
 */
function buildSellasistCard(senderEmail, sellasistData)
{
    // Utworzenie nagłówka karty
    var cardHeader = CardService.newCardHeader()
        .setTitle("Sellasist Gmail Add-on")
        .setSubtitle("Klient: " + senderEmail);

    var section = CardService.newCardSection();

    // Obsługa błędów i wyświetlanie danych
    if (sellasistData.error)
    {
        section.addWidget(
            CardService.newTextParagraph().setText("Wystąpił błąd: " + sellasistData.error)
        );
    } else
    {
        var orders = sellasistData.orders || [];
        if (orders.length === 0)
        {
            section.addWidget(
                CardService.newTextParagraph().setText("Brak zamówień dla " + senderEmail)
            );
        } else
        {
            // Dla każdego zamówienia dodaj tekst i przycisk
            orders.forEach(function (order)
            {
                if (!order) return;

                var createdAt = order.date || "brak daty";
                var status = order.status ? order.status.name : "brak statusu";

                // Dodaj informacje o zamówieniu
                section.addWidget(
                    CardService.newTextParagraph().setText(
                        "Zamówienie #" + order.id + " (" + status + "), data: " + createdAt
                    )
                );

                // Dodaj przycisk do zamówienia
                var orderUrl = "https://" + WEBSITE_NAME + ".sellasist.pl/admin/orders/edit/" + order.id;
                section.addWidget(
                    CardService.newTextButton()
                        .setText("Zobacz zamówienie")
                        .setOpenLink(CardService.newOpenLink()
                            .setUrl(orderUrl)
                            .setOpenAs(CardService.OpenAs.FULL_SIZE)
                        )
                );
            });
        }
    }

    // Zwrócenie skonstruowanej karty
    return CardService.newCardBuilder()
        .setHeader(cardHeader)
        .addSection(section)
        .build();
}