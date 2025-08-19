import cfg.ConfigManager;
import org.jsoup.Jsoup;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimeUtility;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Properties;

public class QQEmailReader {
    private String IMAP_HOST;
    private int IMAP_PORT;
    private String username;
    private String password;
    private String downloadDirectory;

    private final ConfigManager configManager;

    public QQEmailReader() {
        configManager = new ConfigManager();
        loadConfiguration();
    }

    private void loadConfiguration() {
        this.IMAP_HOST = configManager.getString("mail.host", "imap.qq.com");
        this.IMAP_PORT = configManager.getInt("mail.port", 993);
        this.username = configManager.getString("mail.username", null);
        this.password = configManager.getString("mail.password", null);
        this.downloadDirectory = configManager.getString("mail.download.directory", "downloads");

        if (username == null || password == null) {
            System.err.println("警告: 邮箱用户名或密码未在配置文件中设置!");
        }
    }

    public void readRecentEmails() {
        try {
            Properties properties = new Properties();
            properties.put("mail.store.protocol", "imaps");
            properties.put("mail.imaps.host", IMAP_HOST);
            properties.put("mail.imaps.port", IMAP_PORT);
            properties.put("mail.imaps.ssl.enable", "true");

            Session session = Session.getDefaultInstance(properties);
            Store store = session.getStore("imaps");

            store.connect(IMAP_HOST, username, password);

            Folder inbox = store.getFolder("INBOX");
            inbox.open(Folder.READ_ONLY);

            int messageCount = inbox.getMessageCount();
            if (messageCount == 0) {
                System.out.println("收件箱中没有邮件");
                return;
            }

            // 只获取最新的一封邮件
            Message[] messages = inbox.getMessages(messageCount, messageCount);

            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

            System.out.println("=== 最新邮件详细信息 ===\n");

            Message message = messages[0];
            System.out.println("==================== 最新邮件 ====================");

            // 基本信息
            System.out.println("消息ID: " + message.getHeader("Message-ID")[0]);
            System.out.println("发件人: " + getFromAddress(message));
            System.out.println("收件人: " + getRecipients(message, Message.RecipientType.TO));
            System.out.println("抄送: " + getRecipients(message, Message.RecipientType.CC));
            System.out.println("密送: " + getRecipients(message, Message.RecipientType.BCC));
            System.out.println("主题: " + (message.getSubject() != null ? message.getSubject() : "无主题"));

            // 时间信息
            if (message.getSentDate() != null) {
                System.out.println("发送时间: " + dateFormat.format(message.getSentDate()));
            }
            if (message.getReceivedDate() != null) {
                System.out.println("接收时间: " + dateFormat.format(message.getReceivedDate()));
            }

            // 邮件状态
            System.out.println("是否已读: " + (message.isSet(Flags.Flag.SEEN) ? "是" : "否"));
            System.out.println("是否重要: " + (message.isSet(Flags.Flag.FLAGGED) ? "是" : "否"));
            System.out.println("邮件大小: " + formatFileSize(message.getSize()));

            // 邮件头信息
            String[] contentType = message.getHeader("Content-Type");
            if (contentType != null && contentType.length > 0) {
                System.out.println("内容类型: " + contentType[0]);
            }

            // 优先级
            String[] priority = message.getHeader("X-Priority");
            if (priority != null && priority.length > 0) {
                System.out.println("优先级: " + getPriorityText(priority[0]));
            }

            // 邮件内容
            System.out.println("\n--- 邮件内容 ---");
            String content = getDetailedContent(message);
            System.out.println(content);

            // 附件信息和下载
            if (hasAttachments(message)) {
                System.out.println("\n--- 附件信息 ---");
                getAttachmentInfo(message);
            }

            System.out.println("\n" + "=".repeat(60) + "\n");

            inbox.close(false);
            store.close();

        } catch (Exception e) {
            System.err.println("读取邮件时发生错误: " + e.getMessage());
            System.out.println(e.getMessage());
        }
    }

    private String getFromAddress(Message message) throws MessagingException {
        Address[] addresses = message.getFrom();
        if (addresses != null && addresses.length > 0) {
            InternetAddress addr = (InternetAddress) addresses[0];
            return addr.getPersonal() != null ?
                    addr.getPersonal() + " <" + addr.getAddress() + ">" :
                    addr.getAddress();
        }
        return "未知发件人";
    }

    private String getRecipients(Message message, Message.RecipientType type) throws MessagingException {
        Address[] recipients = message.getRecipients(type);
        if (recipients == null || recipients.length == 0) {
            return "无";
        }

        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < recipients.length; i++) {
            InternetAddress addr = (InternetAddress) recipients[i];
            if (i > 0) sb.append(", ");
            sb.append(addr.getPersonal() != null ?
                    addr.getPersonal() + " <" + addr.getAddress() + ">" :
                    addr.getAddress());
        }
        return sb.toString();
    }

    private String cleanText(String text) {
        if (text == null) return "";

        // 替换HTML实体
        text = text.replace("&nbsp;", " ");
        text = text.replace("&amp;", "&");
        text = text.replace("&lt;", "<");
        text = text.replace("&gt;", ">");
        text = text.replace("&quot;", "\"");
        text = text.replace("&apos;", "'");
        text = text.replace("&hellip;", "...");
        text = text.replace("&mdash;", "—");
        text = text.replace("&ndash;", "–");

        // 清理多余的空白字符，但保留换行结构
        text = text.replaceAll("[ \\t]+", " "); // 多个空格/制表符替换为单个空格
        text = text.replaceAll("\\n\\s*\\n", "\n\n"); // 多个换行保留为双换行
        text = text.trim();

        return text;
    }

    private String getDetailedContent(Message message) throws Exception {
        if (message.isMimeType("text/plain")) {
            return cleanText(message.getContent().toString());
        } else if (message.isMimeType("text/html")) {
            String htmlContent = message.getContent().toString();
            String plainText = Jsoup.parse(htmlContent).text();
            return cleanText(plainText);
        } else if (message.isMimeType("multipart/*")) {
            MimeMultipart multipart = (MimeMultipart) message.getContent();
            return getMultipartContent(multipart);
        } else {
            return "【未知格式】无法解析内容";
        }
    }

    private String getMultipartContent(MimeMultipart multipart) throws Exception {
        StringBuilder content = new StringBuilder();
        int count = multipart.getCount();
        boolean hasProcessedText = false; // 标记是否已经处理过文本内容

        for (int i = 0; i < count; i++) {
            MimeBodyPart bodyPart = (MimeBodyPart) multipart.getBodyPart(i);
            String disposition = bodyPart.getDisposition();

            if (disposition == null || disposition.equalsIgnoreCase(Part.INLINE)) {
                // 只处理第一个找到的文本内容
                if (!hasProcessedText) {
                    if (bodyPart.isMimeType("text/plain")) {
                        String textContent = bodyPart.getContent().toString();
                        String cleanedText = cleanText(textContent);
                        content.append(cleanedText);
                        hasProcessedText = true;
                    } else if (bodyPart.isMimeType("text/html")) {
                        String htmlContent = bodyPart.getContent().toString();
                        String plainText = Jsoup.parse(htmlContent).text();
                        String cleanedText = cleanText(plainText);
                        content.append(cleanedText);
                        hasProcessedText = true;
                    } else if (bodyPart.isMimeType("multipart/*")) {
                        String nestedContent = getMultipartContent((MimeMultipart) bodyPart.getContent());
                        if (!nestedContent.isEmpty()) {
                            content.append(nestedContent);
                            hasProcessedText = true;
                        }
                    }
                }
            }
        }

        return content.toString();
    }

    private boolean hasAttachments(Message message) throws Exception {
        if (message.isMimeType("multipart/*")) {
            MimeMultipart multipart = (MimeMultipart) message.getContent();
            return hasAttachmentsInMultipart(multipart);
        }
        return false;
    }

    private boolean hasAttachmentsInMultipart(MimeMultipart multipart) throws Exception {
        int count = multipart.getCount();
        for (int i = 0; i < count; i++) {
            MimeBodyPart bodyPart = (MimeBodyPart) multipart.getBodyPart(i);
            String disposition = bodyPart.getDisposition();

            if (disposition != null && disposition.equalsIgnoreCase(Part.ATTACHMENT)) {
                return true;
            }

            // 检查是否有文件名（某些附件可能没有disposition）
            String fileName = bodyPart.getFileName();
            if (fileName != null && !bodyPart.isMimeType("text/*")) {
                return true;
            }

            // 递归检查嵌套的multipart
            if (bodyPart.isMimeType("multipart/*")) {
                if (hasAttachmentsInMultipart((MimeMultipart) bodyPart.getContent())) {
                    return true;
                }
            }
        }
        return false;
    }

    private void getAttachmentInfo(Message message) throws Exception {
        if (message.isMimeType("multipart/*")) {
            MimeMultipart multipart = (MimeMultipart) message.getContent();

            // 创建下载目录
            Path downloadPath = Paths.get(downloadDirectory);
            if (!Files.exists(downloadPath)) {
                Files.createDirectories(downloadPath);
                System.out.println("已创建下载目录: " + downloadPath.toAbsolutePath());
            }

            processAttachments(multipart, 0);
        }
    }

    private int processAttachments(MimeMultipart multipart, int attachmentCount) throws Exception {
        int count = multipart.getCount();

        for (int i = 0; i < count; i++) {
            MimeBodyPart bodyPart = (MimeBodyPart) multipart.getBodyPart(i);
            String disposition = bodyPart.getDisposition();
            String fileName = bodyPart.getFileName();

            // 判断是否为附件
            boolean isAttachment = false;
            if (disposition != null && disposition.equalsIgnoreCase(Part.ATTACHMENT)) {
                isAttachment = true;
            } else if (fileName != null && !bodyPart.isMimeType("text/*")) {
                isAttachment = true;
            }

            if (isAttachment) {
                attachmentCount++;

                if (fileName != null) {
                    try {
                        // 解码文件名（处理编码问题）
                        fileName = MimeUtility.decodeText(fileName);
                    } catch (Exception e) {
                        System.err.println("文件名解码失败，使用原文件名: " + fileName);
                    }

                    System.out.println("附件 " + attachmentCount + ": " + fileName +
                            " (大小: " + formatFileSize(bodyPart.getSize()) + ")");

                    // 下载附件
                    downloadAttachment(bodyPart, fileName);
                } else {
                    String defaultName = getString(attachmentCount, bodyPart);

                    System.out.println("附件 " + attachmentCount + ": " + defaultName +
                            " (大小: " + formatFileSize(bodyPart.getSize()) + ")");

                    // 下载未命名附件
                    downloadAttachment(bodyPart, defaultName);
                }
            } else if (bodyPart.isMimeType("multipart/*")) {
                // 递归处理嵌套的multipart
                attachmentCount = processAttachments((MimeMultipart) bodyPart.getContent(), attachmentCount);
            }
        }

        return attachmentCount;
    }

    private static String getString(int attachmentCount, MimeBodyPart bodyPart) throws MessagingException {
        String defaultName = "attachment_" + attachmentCount;
        // 尝试根据Content-Type推断扩展名
        String contentType = bodyPart.getContentType();
        if (contentType.contains("pdf")) {
            defaultName += ".pdf";
        } else if (contentType.contains("image")) {
            if (contentType.contains("jpeg") || contentType.contains("jpg")) {
                defaultName += ".jpg";
            } else if (contentType.contains("png")) {
                defaultName += ".png";
            }
        } else if (contentType.contains("application/msword")) {
            defaultName += ".doc";
        } else if (contentType.contains("application/vnd.openxmlformats")) {
            defaultName += ".docx";
        }
        return defaultName;
    }

    private void downloadAttachment(MimeBodyPart bodyPart, String originalFileName) {
        try {
            // 创建文件路径
            Path filePath = Paths.get(downloadDirectory, originalFileName);

            // 如果文件已存在，添加序号
            String fileName;
            int counter = 1;
            while (Files.exists(filePath)) {
                String nameWithoutExt;
                String extension = "";

                int lastDot = originalFileName.lastIndexOf('.');
                if (lastDot > 0) {
                    nameWithoutExt = originalFileName.substring(0, lastDot);
                    extension = originalFileName.substring(lastDot);
                } else {
                    nameWithoutExt = originalFileName;
                }

                fileName = nameWithoutExt + "_(" + counter + ")" + extension;
                filePath = Paths.get(downloadDirectory, fileName);
                counter++;
            }

            // 下载文件
            try (InputStream inputStream = bodyPart.getInputStream();
                 FileOutputStream outputStream = new FileOutputStream(filePath.toFile())) {

                byte[] buffer = new byte[8192];
                int bytesRead;
                long totalBytes = 0;

                while ((bytesRead = inputStream.read(buffer)) != -1) {
                    outputStream.write(buffer, 0, bytesRead);
                    totalBytes += bytesRead;
                }

                System.out.println("  -> 已下载到: " + filePath.toAbsolutePath() +
                        " (实际大小: " + formatFileSize(totalBytes) + ")");
            }

        } catch (Exception e) {
            System.err.println("下载附件失败: " + originalFileName + " - " + e.getMessage());
        }
    }

    private String formatFileSize(long size) {
        if (size <= 0) return "未知";
        if (size < 1024) return size + " B";
        if (size < 1024 * 1024) return String.format("%.1f KB", size / 1024.0);
        if (size < 1024 * 1024 * 1024) return String.format("%.1f MB", size / (1024.0 * 1024));
        return String.format("%.1f GB", size / (1024.0 * 1024 * 1024));
    }

    private String getPriorityText(String priority) {
        return switch (priority) {
            case "1" -> "最高";
            case "2" -> "高";
            case "4" -> "低";
            case "5" -> "最低";
            default -> "普通";
        };
    }

    public static void main(String[] args) {
        QQEmailReader reader = new QQEmailReader();

        // 可选：设置自定义下载目录
        // reader.setDownloadDirectory("D:/EmailAttachments");

        reader.readRecentEmails();
    }
}