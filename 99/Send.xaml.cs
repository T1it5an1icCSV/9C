using System.Net.Mail;
using System.Net;
using System.IO;
using System.Windows;
using MaterialDesignThemes.Wpf;
using System;

namespace _99
{
    public partial class Send : Window
    {
        private string _filePath;
        public Send(string filePath)
        {
            InitializeComponent();
            _filePath = filePath;
        }

        private void Send_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MailMessage mail = new MailMessage(log.Text, For.Text, Theme.Text, "письмо");

                if (File.Exists(_filePath))
                {
                    mail.Attachments.Add(new Attachment(_filePath));
                }

                else
                {
                    MessageBox.Show("Выбран неверный файл!");
                }

                SmtpClient smtp = new SmtpClient();

                if (log.Text.Contains("@mail.ru"))
                {
                    smtp.Host = "smtp.mail.ru";
                    smtp.Port = 587;
                }
                else if (log.Text.Contains("@gmail.com"))
                {
                    smtp.Host = "smtp.gmail.com";
                    smtp.Port = 587;
                }
                else
                {
                    MessageBox.Show("Неверный хост или домен");
                }

                smtp.EnableSsl = true;
                smtp.Credentials = new NetworkCredential(log.Text, pass.Text);

                try
                {
                    smtp.Send(mail);
                    MessageBox.Show("Письмо успешно отправлено!");
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при отправке письма: " + ex);
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Вы не ввели данные для отправки!");
            }
        }

        private void Exit_Button_Click(object sender, RoutedEventArgs e)
        {
            var window = GetWindow(this);

            if (window != null)
            {
                window.Close();
            }
        }
    }
}