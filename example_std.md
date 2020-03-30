# Pstxy Example & Download

## Download
[Download Pstxy Std](Coming soon...)

(If you want to buy Pstxy Plus, please contact: [pantilesoft@outlook.com](mailto:pantilesoft@outlook.com))

## Example
```csharp
class Program
{
    static void Main(string[] args)
    {
        // .pst or .ost file
        var filename = @"d:\pst\sample.pst";
        using (PstFile pstFile = new PstFile(filename))
        {
            foreach (var folder in pstFile.Folders)
            {
                DisplayFolder(folder);
            }
        }

        Console.WriteLine("Done.");
        Console.ReadLine();
    }

    private static void DisplayFolder(Folder folder)
    {
        Console.WriteLine(String.Format("Folder: {0}, Total: {1}, Unread: {2}",
            folder.Name, folder.ContentCount, folder.ContentUnreadCount));

        foreach (var mail in folder.Mails)
        {
            DisplayMail(mail);
        }

        foreach (var subFolder in folder.Folders)
        {
            DisplayFolder(subFolder);
        }
    }

    private static void DisplayMail(Mail mail)
    {
        Console.WriteLine(String.Format("Mail: '{0}' [{1}/{2}]", mail.Subject, mail.SenderName, mail.SenderEmail));
        Console.WriteLine(String.Format("Delivery at:{0}, Create at:{1}", mail.DeliveryTime, mail.CreationTime));

        foreach (Recipient rcpt in mail.Recipients)
        {
            Console.WriteLine(String.Format("Recipient: {0}/{1}", rcpt.Name, rcpt.Email));
        }

        if (mail.Content != null)
        {
            // text mail content
            Console.WriteLine(mail.Content);
        }
        if (mail.ContentHtml != null)
        {
            // html mail content
            Console.WriteLine(mail.ContentHtml);
        }
        if (mail.ContentRtf != null)
        {
            // rtf mail content
            Console.WriteLine(mail.ContentRtf);
        }
    }
}
```
