using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Newtonsoft.Json;
using System.Reflection;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using System.Web.Http.Cors;

namespace WordService.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class ValuesController : ApiController
    { 

        // POST api/values
        public IHttpActionResult Post([FromBody] Doc doc)
        {
            try
            { 
                //doc.json = "{\"Resume\": \"333\"}";
                //doc.base64 = "UEsDBBQABgAIAAAAIQDfpNJsWgEAACAFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC0lMtuwjAQRfeV+g+Rt1Vi6KKqKgKLPpYtUukHGHsCVv2Sx7z+vhMCUVUBkQpsIiUz994zVsaD0dqabAkRtXcl6xc9loGTXmk3K9nX5C1/ZBkm4ZQw3kHJNoBsNLy9GUw2ATAjtcOSzVMKT5yjnIMVWPgAjiqVj1Ykeo0zHoT8FjPg973eA5feJXApT7UHGw5eoBILk7LXNX1uSCIYZNlz01hnlUyEYLQUiep86dSflHyXUJBy24NzHfCOGhg/mFBXjgfsdB90NFEryMYipndhqYuvfFRcebmwpCxO2xzg9FWlJbT62i1ELwGRztyaoq1Yod2e/ygHpo0BvDxF49sdDymR4BoAO+dOhBVMP69G8cu8E6Si3ImYGrg8RmvdCZFoA6F59s/m2NqciqTOcfQBaaPjP8ber2ytzmngADHp039dm0jWZ88H9W2gQB3I5tv7bfgDAAD//wMAUEsDBBQABgAIAAAAIQAekRq37wAAAE4CAAALAAgCX3JlbHMvLnJlbHMgogQCKKAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArJLBasMwDEDvg/2D0b1R2sEYo04vY9DbGNkHCFtJTBPb2GrX/v082NgCXelhR8vS05PQenOcRnXglF3wGpZVDYq9Cdb5XsNb+7x4AJWFvKUxeNZw4gyb5vZm/cojSSnKg4tZFYrPGgaR+IiYzcAT5SpE9uWnC2kiKc/UYySzo55xVdf3mH4zoJkx1dZqSFt7B6o9Rb6GHbrOGX4KZj+xlzMtkI/C3rJdxFTqk7gyjWop9SwabDAvJZyRYqwKGvC80ep6o7+nxYmFLAmhCYkv+3xmXBJa/ueK5hk/Nu8hWbRf4W8bnF1B8wEAAP//AwBQSwMEFAAGAAgAAAAhANZks1H0AAAAMQMAABwACAF3b3JkL19yZWxzL2RvY3VtZW50LnhtbC5yZWxzIKIEASigAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArJLLasMwEEX3hf6DmH0tO31QQuRsSiHb1v0ARR4/qCwJzfThv69ISevQYLrwcq6Yc8+ANtvPwYp3jNR7p6DIchDojK971yp4qR6v7kEQa1dr6x0qGJFgW15ebJ7Qak5L1PWBRKI4UtAxh7WUZDocNGU+oEsvjY+D5jTGVgZtXnWLcpXndzJOGVCeMMWuVhB39TWIagz4H7Zvmt7ggzdvAzo+UyE/cP+MzOk4SlgdW2QFkzBLRJDnRVZLitAfi2Myp1AsqsCjxanAYZ6rv12yntMu/rYfxu+wmHO4WdKh8Y4rvbcTj5/oKCFPPnr5BQAA//8DAFBLAwQUAAYACAAAACEAqEHum+cCAABPCgAAEQAAAHdvcmQvZG9jdW1lbnQueG1spJZZb9swDIDfB+w/GH5v5SOHYzQp2mUt+jCgaLfnQZHlA7EOSErc7NePsuM4g7vCcRFAlkTxEymRjG5u31jp7KnSheBL17/2XIdyIpKCZ0v318+Hq8h1tME8waXgdOkeqHZvV1+/3FRxIsiOUW4cQHAdV5Is3dwYGSOkSU4Z1tesIEpokZprIhgSaVoQiiqhEhR4vlf3pBKEag37fcN8j7V7xJG3YbRE4QqULXCCSI6VoW8dw78YMkULFPVBwQgQeBj4fVR4MWqGrFU90GQUCKzqkabjSO84NxtHCvqk+ThS2CdF40i9cGL9ABeSchCmQjFsYKgyxLDa7uQVgCU2xaYoC3MApjdrMbjg2xEWgdaJwMLkYsIcMZHQMkxaili6O8Xjo/7VSd+aHjf6x0+roYb436isj8Wh9hwpWsJZCK7zQp4ynI2lgTBvIfuPnNizsl1XSX9guvyvPK2bo+yAQ8w/nj8rG8s/JvregBuxiJPGEBP+3bO1hEEUdhuPOpqzw/UHFpAWEPQAM1IMDOmW0Zwm+AOaZxxNL8NMW4w+sC7VK5l9LloeldjJjlZ8jvbU5X5l/4UvYB2j7jwT9OeMec2xhJLASPyUcaHwpgSLIIYcCAOnvgHbwq04NuncFTwVNiI52K8EySSWWOEnuO3wIXiYBzN4YdhZKLTGzvprb/397h4eIlUMz5LkZel6XjD3wrvwNLWmKd6Vxkq8cB5M1/UuyjZm9UI1FAvH8pwS298NsvO2rZdshNjaKv1qoLwD0kZevR3HDFz5/SjuMdm66Hztd56cVqIaJa1YU2Ke1TuW1t5mr39ABFnq+wtb/6s4h/4sCqMGLrMf2CobAcXEn0wan4ssB8/8yKuHG2GMYJ24pOmZNKc4oVCWo6lvh6kQxg4Xi8AOs52ph7XJVUxEqWFWS0zAzUkwbabhHfeo7G3F5iBBUBYcHnl2K+g8F4aA0aEfHN1uPK67za2i7iW4+gsAAP//AwBQSwMEFAAGAAgAAAAhADoFzBnhBgAAziAAABUAAAB3b3JkL3RoZW1lL3RoZW1lMS54bWzsWVtrG0cUfi/0Pyz7rui2q4uJHKSVFDexExMrKXkcr0a7Y83uiJmRHRECIXkqhUIhLXlooPSlD6U00EBD+9D/UpeENP0RnZmVtDvSLE5iG0Kxbay5fOfMN+ecOXO0e/nKvQhbh5AyROKWXb5Usi0Y+2SI4qBl3x70Cw3bYhzEQ4BJDFv2DDL7yuann1wGGzyEEbSEfMw2QMsOOZ9sFIvMF8OAXSITGIu5EaER4KJLg+KQgiOhN8LFSqlUK0YAxbYVg0iovTkaIR9af7/8480PT/96+KX4szcXa/Sw+BdzJgd8TPfkClATVNjhuCw/2Ix5mFqHALdssdyQHA3gPW5bGDAuJlp2Sf3Yxc3LxaUQ5jmyGbm++pnLzQWG44qSo8H+UtBxXKfWXupXAMzXcb16r9arLfUpAPB9sdOEi66zXvGcOTYDSpoG3d16t1rW8Bn91TV825W/Gl6Bkqazhu/3vdSGGVDSdNfwbqfZ6er6FShp1tbw9VK769Q1vAKFGMXjNXTJrVW9xW6XkBHBW0Z403X69cocnqKKmehK5GOeF2sROCC0LwDKuYCj2OKzCRwBX+Be//zF69//tLZREIq4m4CYMDFaqpT6par4L38d1VIOBRsQZISTIZ+tDUk6FvMpmvCWfU1otTOQVy9fHj96cfzot+PHj48f/TJfe11uC8RBVu7tj1//++yh9c+v37998o0Zz7J4bWtGONdoffv89Yvnr55+9eanJwZ4m4L9LHyAIsisG/DIukUisUHDAnCfvp/EIAQoK9GOAwZiIGUM6B4PNfSNGcDAgOtA3Y53qMgWJuDV6YFGeC+kU44MwOthpAF3CMEdQo17ui7XylphGgfmxek0i7sFwKFpbW/Fy73pRIQ9Mqn0QqjR3MXC5SCAMeSWnCNjCA1idxHS7LqDfEoYGXHrLrI6ABlNMkD7WjSlQlsoEn6ZmQgKf2u22bljdQg2qe/CQx0pzgbAJpUQa2a8CqYcREbGIMJZ5DbgoYnk3oz6msEZF54OICZWbwgZM8ncpDON7nUg0pbR7Tt4FulIytHYhNwGhGSRXTL2QhBNjJxRHGaxn7GxCFFg7RJuJEH0EyL7wg8gznX3HQQ1d598tm+LNGQOEDkzpaYjAYl+Hmd4BKBJeZtGWoptU2SMjs400EJ7G0IMjsAQQuv2ZyY8mWg2T0lfC0VW2YIm21wDeqzKfgwZtFRtY3AsYlrI7sGA5PDZma0knhmII0DzNN8Y6yHT26fiMJriFftjLZUiKg+tmcRNFmn7y9W6GwItrGSfmeN1RjX/vcsZEzIHHyAD31tGJPZ3ts0AYG2BNGAGAFnbpnQrRDT3pyLyOCmxqVFupB/a1A3FlZonQvFJBdBK6eOeX+kjCoxX3z0zYM+m3DEDT1Po5OWS1fImD7da1HiEDtHHX9N0wTTeheIaMUAvSpqLkuZ/X9LkneeLQuaikLkoZMwi51DIpLWLegC0eMyjtES5z3xGCOM9PsNwm6mqh4mzP+yLQdVRQstHTJNQNOfLabiAAtW2KOGfIx7uhWAilimrFQI2Vx0wa0KYKJzUsFG3nMDTaIcMk9FyefFUUwgAno6LwmsxLqo0nozW6unju6V61QvUY9YFASn7PiQyi+kkqgYS9cXgCSTUzs6ERdPAoiHV57JQH3OviMvJAvK5uOskjES4iZAeSj8l8gvvnrmn84ypb7ti2F5Tcj0bT2skMuGmk8iEYSguj9XhM/Z1M3WpRk+aYp1GvXEevpZJZCU34FjvWUfizFVdocYHk5Y9Et+YRDOaCH1MZiqAg7hl+3xu6A/JLBPKeBewMIGpqWT/EeKQWhhFItazbsBxyq1cqcs9fqTkmqWPz3LqI+tkOBpBn+eMpF0xlygxzp4SLDtkKkjvhcMjax9P6S0gDOXWy9KAQ8T40ppDRDPBnVpxJV3Nj6L2tiU9ogBPQjC/UbLJPIGr9pJOZh+K6equ9P58M/uBdNKpb92TheREJmnmXCDy1jTnj/O75DOs0ryvsUpS92quay5yXd4tcfoLIUMtXUyjJhkbqKWjOrUzLAgyyy1DM++OOOvbYDVq5QWxqCtVb+21Ntk/EJHfFdXqFHOmqIpvLRR4ixeSSSZQo4vsco9bU4pa9v2S23a8iusVSg23V3CqTqnQcNvVQtt1q+WeWy51O5UHwig8jMpusnZffNnHs/nLezW+9gI/WpTal3wSFYmqg4tKWL3AL1dML/AHct62kLDM/Vql36w2O7VCs9ruF5xup1FoerVOoVvz6t1+13Mbzf4D2zpUYKdd9Zxar1GolT2v4NRKkn6jWag7lUrbqbcbPaf9YG5rsfPF58K8itfmfwAAAP//AwBQSwMEFAAGAAgAAAAhAAFV0fasBAAAngwAABEAAAB3b3JkL3NldHRpbmdzLnhtbLRX227bOBB9X2D/wdDzOpZlyXaMOoUvUZMiaYso2T5TEm1xQ5ECSdlxiv33Hd5i51YkXfSlpefMnBkO56J8+HhX084GC0k4mwb9ozDoYFbwkrD1NLi5TrvjoCMVYiWinOFpsMMy+Hjy5x8fthOJlQI12QEKJid1MQ0qpZpJryeLCtdIHvEGMwBXXNRIwU+x7tVI3LZNt+B1gxTJCSVq14vCcBg4Gj4NWsEmjqJbk0JwyVdKm0z4akUK7P7zFuItfq3JkhdtjZkyHnsCU4iBM1mRRnq2+lfZAKw8yeZnl9jU1Ott++EbrrvlonyweEt42qARvMBSwgPV1AdI2N5x/IzowfcR+HZXNFRg3g/N6TDy5H0E0TOCYUHK93EMHUcPLA94JH4fTeJp5K7Gd55I0rek1kIXJBdI2MJ1ea2LyfmacYFyCuFAfjuQoo6JTv+rIz6BprnnvO5sJw0WBVTONDg+DnpankNE0IVL/oWrrBWCt6w8wwhkr8Ip58rBJV6hlqprlGeKN8C/QXCbOAoteSnQFgrhkyDl31goUiCaNagAkVftJ0OnSmRD0e6MC3LPmUJ0ubc9hTmx8xae2up72te0I6tdVEigAqJ27hfgQnDqtfRUEFC031pWqNb0prMz40KfJBjilIubC5sXRBErcAZcFM93Cnqyze3pOylVZWPUWbvAaIPnqLiVFMlqpqeZAVt6LRAx+bACo31618DMyyqyUldYQYcaCJX/tFJdEIbPMFlX6pxd6+e2PBKnpxdox1t1EHJmZyRckKEa2xs+zL1LXsIQA1NB3l7A2sA92WFunjrikH14BGwCzNSOQtKYysg9nrHyM9yCAKPN8K9H8LMAMNOev0IHXe8anGIEWYTd8XucmTdLKWkuCfSGOGclNNdvc0ZWKyzAAUEKX0LbEcG3Js+2YX+XX6iw76AMA2wAJVvczrlSvD7bNRXk+v+9pGnm3mGfwQdBKf3hCibNg2oYx0k6c5FqdI+Eg1GULF9ColE4mA1eQuJoOItdKT9BXvUziwfh6EU/yygZham7jbtDPdGr+ZvwJ90IndpaLFCdC4I6l3p597RGLm7nhHk8xzD88SGStbkHu10LyBpRmsKTeMCkszazcYlX5kwvkVjveZ2GeFEK0/zzA5deE1h8gonfWHQrUGML3Kv049hZEgbTqfZy2eaZt2Kwrg4gWB9fN8LkaZ+e7URBwZhBcYH2uwGz7k2mSwUjqWaSoGlwX3UXX1ytUpHpOsOXqGlsuebr/jSgekD2tZmCXyV89pkf+TpyWGSwyGLmByr0ZUHbHfayyMsO9AZeNtjLYi+L97LEy5K9bOhlQy2rYEAJStgtdI4/avmKU8q3uDzb489ENgmyQg1e2iUMFcetwG1l2dlM8B3selwSBV/TDSlrdAfPFkZm6TptanbHI12NaeXmMUOJFHKzovfI2FT9k1j0x0FBoEKzXZ3vV+2RDZwSCXOmga2suPDYXwbrx7AKi3NoLjjZOjsdjcNBPLZwYra5MqMI3v0Kr+ZI4tJh3jSxpj/S/nDeT2ejbni8WHTjxXjcHc+Gs+5svBiEyTxNBkn6r+tb/4fFyX8AAAD//wMAUEsDBBQABgAIAAAAIQDv2sPmggEAAOsCAAARAAgBZG9jUHJvcHMvY29yZS54bWwgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB8kt9OwjAUxu9NfIel96MdBALLGIkariQxEaPxrrYHqGxd0xYGL4G33pj4dPoadhsbLhLvzp/v/Hr6tdFklybeFrQRmRyjoEOQB5JlXMjlGD3Mp/4QecZSyWmSSRijPRg0iS8vIqZClmm405kCbQUYz5GkCZkao5W1KsTYsBWk1HScQrrmItMptS7VS6woW9Ml4C4hA5yCpZxaigugrxoiOiI5a5Bqo5MSwBmGBFKQ1uCgE+CT1oJOzdmBsvNLmQq7V3BWWjcb9c6IRpjneSfvlVK3f4CfZrf35VV9IQuvGKA44iy0wiYQR/gUushsXl6B2arcJC5mGqjNdPz1cfC+3g/fb5+lpC4Xhq9hn2eaGzfcypyMg2FaKOuesUK3Ck6dUGNn7l0XAvjVvnXK324xoGEril8R90tFk0ZHi6vNgHvOmrAysu489q5v5lMUd0kw9MnAJ/05GYa9UUjIc7Fca/4ETI8L/E8c+aTrdwdz4nBBm1gDKn/a3zP+AQAA//8DAFBLAwQUAAYACAAAACEA8gQHt/UBAACkBQAAEgAAAHdvcmQvZm9udFRhYmxlLnhtbNyS0W6bMBSG7yftHSzfNxgINEMlVZsWadLUi6mTdus4BqxhG9lOaB5hD7MX2M0ep6/Rg4G2ajotuV2iEPMf+zv2J19cPsgG7bixQqschzOCEVdMb4SqcvztvjhbYGQdVRvaaMVzvOcWXy4/frjoslIrZxGsVzaTLMe1c20WBJbVXFI70y1XUCy1kdTBq6kCSc2PbXvGtGypE2vRCLcPIkJSPGLMMRRdloLxG822kivn1weGN0DUytaitROtO4bWabNpjWbcWjizbAaepEI9Y8L5AUgKZrTVpZvBYcYdeRQsD4kfyeYFkJwGiA4AKROb0xjpyAhg5SuO5adhkglj95I/YCRZ9rlS2tB1AyRQg+B0yIP7Z99sOd4N1GWKSph1w1X1XVDlK7Rxd5BCcUebHD/++vn4+w8O+lJLlbY8nEqkV5kSQmL4H7/DRFZTY3nfwE9cpENcUima/ZTSrdMjVzhWT/GOGtFvfihZUUFha9ckx1fQikTXBR6SMMfxYlWcr4qrMYlgT/4TpmMSTwkhfcI8B17m8PMc5jnPc6BnMMg5kHQvJLfojnfoq5ajq0MhEQiJSQINEhjHZP6ukKHTWyHGc08xctsLuS1eGVlBcr5Irt8aIZ/+YQSkDZzjjQxXA30RVe3+ouN/vh/jwC6fAAAA//8DAFBLAwQUAAYACAAAACEAvdSNvycBAACPAgAAFAAAAHdvcmQvd2ViU2V0dGluZ3MueG1slNLNagIxEADge6HvEHLXrFKlLK5CKZZeSqHtA8TsrIZmMiETu9qnb9xqf/DiXkImyXzJhJktdujEB0S25Cs5GhZSgDdUW7+u5NvrcnArBSfta+3IQyX3wHIxv76atWULqxdIKZ9kkRXPJZpKblIKpVJsNoCahxTA582GIuqUw7hWqOP7NgwMYdDJrqyzaa/GRTGVRyZeolDTWAP3ZLYIPnX5KoLLInne2MAnrb1EaynWIZIB5lwPum8PtfU/zOjmDEJrIjE1aZiLOb6oo3L6qOhm6H6BST9gfAZMja37GdOjoXLmH4ehHzM5MbxH2EmBpnxce4p65bKUv0bk6kQHH8bDZfPcIRSSRfsJS4p3kVqGqA7L2jlqn58ecqD+tdH8CwAA//8DAFBLAwQUAAYACAAAACEACLyHSVgLAACvcQAADwAAAHdvcmQvc3R5bGVzLnhtbLydTXPbOBKG71u1/4Gl0+7Bkb+dpMaZcpxk7drY4xk5mzNEQhbGIKEFKcueX78gSEqQm6DYYK8viUWxH4J48QJo8EO//PqcyuiJ61yo7Hx08G5/FPEsVonIHs5HP+6/7b0fRXnBsoRJlfHz0QvPR79++vvffll9zIsXyfPIALL8Yxqfj+ZFsfg4HufxnKcsf6cWPDNfzpROWWE+6odxyvTjcrEXq3TBCjEVUhQv48P9/dNRjdF9KGo2EzH/ouJlyrPCxo81l4aosnwuFnlDW/WhrZROFlrFPM/NSaey4qVMZGvMwTEApSLWKlez4p05mbpEFmXCD/btX6ncAE5wgEMAOI1FgmOc1oyxiXQ4OcdhThpM/pLy51GUxh+vHzKl2VQakqmayJxdZMHlv+XBPpnGkaj4C5+xpSzy8qO+0/XH+pP975vKijxafWR5LMS9KYwhpsLAry6yXIzMN5zlxUUumPvl13pb+f283LE1Ms4LZ/NnkYjRuDzoI9eZ+fqJyfPRYbUp/2u94aDZclmWq9pW7yVZ9tBs49nej4lbvvPRX/O9y9ty09Qc6nzE9N7kogwc16db/e9UwmL9qdrrVY2Z5mwa96TymPmWz76r+JEnk8J8cT7aLw9lNv64vtNCaeOj89GHD/XGCU/FlUgSnjk7ZnOR8J9znv3IebLZ/vs364V6Q6yWmfn76OzEqijz5OtzzBels8y3GSsr9LYMkOXeS7E5uA3/bwOr67E1fs5Z2b1EB68RtvgoxGEZkTtn285cvjp3uxfqQEdvdaDjtzrQyVsd6PStDnT2Vgd6/1YHspj/54FElvDnyojwMIC6i+NxI5rjMRua4/ESmuOxCprjcQKa42noaI6nHaM5nmaK4BQq9rVCp7EfeVp7N3f3GBHG3T0khHF3jwBh3N0dfhh3d/8ext3dnYdxd/feYdzdnTWeW021omtjs6wY7LKZUkWmCh4V/Hk4jWWGZXMuGl456HFNcpIEmKpnqwfiwbSY2c+7W4g1afh4XpRZWaRm0Uw8LLVJ1YcWnGdPXJqkOWJJYniEQM2LpfbUSEib1nzGNc9iTtmw6aBSZDzKlumUoG0u2AMZi2cJcfU1RJJOYd2g2bKYlyYRBI06ZbFWw4umGFn/8F3kw+uqhESfl1JyItYtTROzrOG5gcUMTw0sZnhmYDHDEwNHM6oqqmlENVXTiCqsphHVW9U+qeqtphHVW00jqreaNrze7kUhbRfvzjoO+q/dXUpVrpIPLsdEPGTMTACGDzf1mml0xzR70Gwxj8pV5Xase87Y43xWyUt0TzGmrUlU83rbRC7NWYtsObxCt2hU5lrziOy15hEZbM0bbrEbM00uJ2hXNPnMZDktWk1rSb1MO2FyWU1oh7uNFcNb2MYA34TOyWzQjiVowbfldLaUk6Ln25RyeME2rOG2et0rkRavRhKUUqr4kaYbvnpZcG3SssfBpG9KSrXiCR1xUmhVtTXX8odWkl6W/5ou5iwXNlfaQvQf6pvr69ENWww+oTvJREaj29e9lAkZ0c0gru5vvkf3alGmmWXF0AA/q6JQKRmzXgn8x08+/SdNAS9MEpy9EJ3tBdHykIVdCoJBpiKphIhkppkiEyRjqOX9m79MFdMJDe1O8+qWloITEScsXVSTDgJvmX5xZfofgtmQ5f2HaVGuC1GZ6p4E5iwb5svpnzwe3tXdqohkZei3ZWHXH+1U10bT4YZPE7Zww6cIVk0zPJTtl+Bkt3DDT3YLR3Wyl5LlufBeQg3mUZ1uw6M+3+HJX81TUunZUtJVYAMkq8EGSFaFSi7TLKc8Y8sjPGHLoz5fwiZjeQRLcpb3Ly0SMjEsjEoJC6OSwcKoNLAwUgGG36HjwIbfpuPAht+rU8GIpgAOjKqdkQ7/RFd5HBhVO7MwqnZmYVTtzMKo2tnRl4jPZmYSTDfEOEiqNucg6QaarODpQmmmX4iQXyV/YAQLpBXtTqtZ+ayDyqqbuAmQ5Rq1JJxsVzgqkX/yKVnRShZluQhWRJmUShGtrW0GHBu5fe/arjD7GMbgItxJFvO5kgnXnnPyx5p8ebJgcb1MDy739Vr2/C4e5kU0ma9X+13M6f7OyCZh3wrbfcC2Oj9tnjxpC7vhiVimTUHhwxSnR/2DbYveCj7eHbyZSWxFnvSMhMc83R25mSVvRZ71jITHfN8z0vp0K7LLD1+YfmxtCGdd7Wed43ka31lXK1oHtx62qyGtI9ua4FlXK9qySnQRx+XVAqhOP8/44/uZxx+PcZGfgrGTn9LbV35El8H+4E+iHNkxnaY93vruCdDv20l0r57z96Wq1u23Ljj1f6jr2kycspxHrZyj/heutnoZfz327m78iN79jh/RuwPyI3r1RN5wVJfkp/Tum/yI3p2UH4HureCIgOutYDyut4LxIb0VpIT0VgNmAX5E7+mAH4E2KkSgjTpgpuBHoIwKwoOMCiloo0IE2qgQgTYqnIDhjArjcUaF8SFGhZQQo0IK2qgQgTYqRKCNChFoo0IE2qiBc3tveJBRIQVtVIhAGxUi0Ea188UBRoXxOKPC+BCjQkqIUSEFbVSIQBsVItBGhQi0USECbVSIQBkVhAcZFVLQRoUItFEhAm3U6lHDcKPCeJxRYXyIUSElxKiQgjYqRKCNChFoo0IE2qgQgTYqRKCMCsKDjAopaKNCBNqoEIE2qr1YOMCoMB5nVBgfYlRICTEqpKCNChFoo0IE2qgQgTYqRKCNChEoo4LwIKNCCtqoEIE2KkR0tc/6EqXvNvsD/Kqn9479/peu6kL94T7K7aKO+qOaUvlZ/Z9F+KzUY9T64OGRzTf6QcRUCmWXqD2X1V2uvSUCdeHzt8vuJ3xc+sCXLtXPQthrpgB+3DcSrKkcdzV5NxIkecddLd2NBLPO467e140Ew+BxV6drfdnclGKGIxDc1c04wQee8K7e2gmHVdzVRzuBsIa7emYnEFZwV3/sBJ5EZef8OvqkZz2dru8vBYSu5ugQzvyErmYJtWq6Y2iMvqL5CX3V8xP6yugnoPT0YvDC+lFohf2oMKmhzbBShxvVT8BKDQlBUgNMuNQQFSw1RIVJDTtGrNSQgJU6vHP2E4KkBphwqSEqWGqICpMaDmVYqSEBKzUkYKUeOCB7MeFSQ1Sw1BAVJjWc3GGlhgSs1JCAlRoSgqQGmHCpISpYaogKkxpkyWipIQErNSRgpYaEIKkBJlxqiAqWGqK6pLarKFtSoxR2wnGTMCcQNyA7gbjO2QkMyJac6MBsySEEZktQq0ZzXLbkiuYn9FXPT+gro5+A0tOLwQvrR6EV9qPCpMZlS21ShxvVT8BKjcuWvFLjsqVOqXHZUqfUuGzJLzUuW2qTGpcttUkd3jn7CUFS47KlTqlx2VKn1LhsyS81LltqkxqXLbVJjcuW2qQeOCB7MeFS47KlTqlx2ZJfaly21CY1LltqkxqXLbVJjcuWvFLjsqVOqXHZUqfUuGzJLzUuW2qTGpcttUmNy5bapMZlS16pcdlSp9S4bKlTaly2dGNCBMEroCYp00VE9764K5bPCzb85YQ/Ms1zJZ94EqFPdbza+s2q8hj2F+LM/oU50fK15c4zRkn12tYaaHe8NiRmf3aqLExU/9RW/WtTtsz15VX796L6DbGVSNSqfOZaK9mE1O3qz7jZMFXFvC6iDRvXR4RljOemkHH9pipfGfdBIT0vobXF2DSuZu9agU21VvttVWpVWk8pi7Ixd5XwwFONlQ185fpQ+3pXwUwxprKqfvPHdZYYwKr+BbCqgMkzq1Dm+0su5Q2r9lYL/66Sz4rq24N9+xaCV99PqxfqeeO17Xm9gPF2YaqP3Y2hesV+fUuAr6oPW6ra3psytJY35Wr+yj/9DwAA//8DAFBLAwQUAAYACAAAACEA3XkYP3ABAADHAgAAEAAIAWRvY1Byb3BzL2FwcC54bWwgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACcUstOwzAQvCPxD1HurZMeAFVbI1SEOPCSmrZny94kFo5t2aaif8+GtCGIGz7tzHpHM2vD7WdnsgOGqJ1d5eW8yDO00iltm1W+rR5mN3kWk7BKGGdxlR8x5rf88gLegvMYksaYkYSNq7xNyS8Zi7LFTsQ5tS11ahc6kQiGhrm61hLvnfzo0Ca2KIorhp8JrUI186NgPiguD+m/osrJ3l/cVUdPehwq7LwRCflLP2nmyqUO2MhC5ZIwle6QF0SPAN5Eg5GXwIYC9i6oyBfAhgLWrQhCJtofL6+BTSDceW+0FIkWy5+1DC66OmWv326zfhzY9ApQgg3Kj6DTsTcxhfCk7WBjKMhWEE0Qvj15GxFspDC4puy8FiYisB8C1q7zwpIcGyvSe49bX7n7fg2nkd/kJONep3bjhey93EzTThqwIRYV2R8djAQ80nME08vTrG1Qne/8bfT72w3/kpdX84LO98LOHMUePwz/AgAA//8DAFBLAQItABQABgAIAAAAIQDfpNJsWgEAACAFAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAB6RGrfvAAAATgIAAAsAAAAAAAAAAAAAAAAAkwMAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhANZks1H0AAAAMQMAABwAAAAAAAAAAAAAAAAAswYAAHdvcmQvX3JlbHMvZG9jdW1lbnQueG1sLnJlbHNQSwECLQAUAAYACAAAACEAqEHum+cCAABPCgAAEQAAAAAAAAAAAAAAAADpCAAAd29yZC9kb2N1bWVudC54bWxQSwECLQAUAAYACAAAACEAOgXMGeEGAADOIAAAFQAAAAAAAAAAAAAAAAD/CwAAd29yZC90aGVtZS90aGVtZTEueG1sUEsBAi0AFAAGAAgAAAAhAAFV0fasBAAAngwAABEAAAAAAAAAAAAAAAAAExMAAHdvcmQvc2V0dGluZ3MueG1sUEsBAi0AFAAGAAgAAAAhAO/aw+aCAQAA6wIAABEAAAAAAAAAAAAAAAAA7hcAAGRvY1Byb3BzL2NvcmUueG1sUEsBAi0AFAAGAAgAAAAhAPIEB7f1AQAApAUAABIAAAAAAAAAAAAAAAAApxoAAHdvcmQvZm9udFRhYmxlLnhtbFBLAQItABQABgAIAAAAIQC91I2/JwEAAI8CAAAUAAAAAAAAAAAAAAAAAMwcAAB3b3JkL3dlYlNldHRpbmdzLnhtbFBLAQItABQABgAIAAAAIQAIvIdJWAsAAK9xAAAPAAAAAAAAAAAAAAAAACUeAAB3b3JkL3N0eWxlcy54bWxQSwECLQAUAAYACAAAACEA3XkYP3ABAADHAgAAEAAAAAAAAAAAAAAAAACqKQAAZG9jUHJvcHMvYXBwLnhtbFBLBQYAAAAACwALAMECAABQLAAAAAA=";

                dynamic jsons = JsonConvert.DeserializeObject(doc.json);

                Dictionary<string, string> dicValues = new Dictionary<string, string>();

                foreach (var item in jsons)
                {
                    dicValues.Add((string)item.Path, (string)item.Value);
                }

                Random rd = new Random();
                var name = rd.Next(1, 999999);
                string oldPath = AppContext.BaseDirectory + name + "old.docx";
                System.IO.File.WriteAllBytes(oldPath, Convert.FromBase64String(doc.base64));

                string newPath = AppContext.BaseDirectory + name + "new.docx";

                //WordReplace(oldPath, newPath, dicValues);

                WordTemplateHelper.WriteToPublicationOfResult(oldPath, newPath, dicValues);


                byte[] newBytes = System.IO.File.ReadAllBytes(newPath);

                try
                {
                    System.IO.File.Delete(oldPath);
                    System.IO.File.Delete(newPath);
                }
                catch { }

                return Ok(new { docBase64 = Convert.ToBase64String(newBytes) });
            }
            catch (Exception ex)
            {
                return Ok(new { exception = ex.Message + "; StackTrace:" + ex.StackTrace });
            }
        } 

        public class Doc
        {
            public string base64 { get; set; }
            public string json { get; set; }
        }

        static void WordReplace(string oldPath, string newPath, Dictionary<string, string> dicValues) {
            Object Nothing = Missing.Value;

            object newDoc = newPath;
            Application wordApp;//Word应用程序变量初始化
            Document wordDoc;

            wordApp = new Application();//创建word应用程序

            object fileName = (oldPath);//模板文件

            wordDoc = wordApp.Documents.Open(ref fileName,
            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            object replace = WdReplace.wdReplaceAll;

            wordApp.Selection.Find.Replacement.ClearFormatting();
            wordApp.Selection.Find.MatchWholeWord = true;
            wordApp.Selection.Find.ClearFormatting();

            foreach (var item in dicValues)
            {
                object FindText = item.Key;
                object Replacement = item.Value;

                if (Replacement.ToString().Length > 110)
                {
                    FindAndReplaceLong(wordApp, FindText, Replacement);
                }
                else
                {
                    FindAndReplace(wordApp, FindText, Replacement);
                }
            }

            wordDoc.SaveAs(newDoc,
            Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing,
            Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
            //关闭wordDoc文档
            wordApp.Documents.Close(ref Nothing, ref Nothing, ref Nothing);
            //关闭wordApp组件对象
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        }

        public static void FindAndReplaceLong(Application wordApp, object findText, object replaceText)
        {
            try {
                int len = replaceText.ToString().Length; //要替换的文字长度
                int cnt = len / 110; //不超过220个字
                string newstr;
                object newStrs;
                if (len < 110) //小于220字直接替换
                {
                    FindAndReplace(wordApp, findText, replaceText);
                }
                else
                {
                    for (int i = 0; i <= cnt; i++)
                    {
                        if (i != cnt)
                            newstr = replaceText.ToString().Substring(i * 110, 110) + findText; //新的替换字符串
                        else
                            newstr = replaceText.ToString().Substring(i * 110, len - i * 110); //最后一段需要替换的文字
                        newStrs = (object)newstr;
                        FindAndReplace(wordApp, findText, newStrs); //进行替换
                    }
                }
            }
            catch (Exception ex) {
                throw ex;
            }
        }

        public static void FindAndReplace(Application wordApp, object findText, object replaceText)
        {
            try { 
                object matchCase = true;
                object matchWholeWord = true;
                object matchWildCards = false;
                object matchSoundsLike = false;
                object matchAllWordForms = false;
                object forward = true;
                object format = false;
                object matchKashida = false;
                object matchDiacritics = false;
                object matchAlefHamza = false;
                object matchControl = false;
                object read_only = false;
                object visible = true;
                object replace = 2;
                object wrap = 1;
                wordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards,
                ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceText,
                ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }
            catch (Exception ex) {
                throw ex;
            }
        }

    }

}
