package com.example.excelapp;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import org.springframework.stereotype.Service;
import org.telegram.telegrambots.bots.TelegramLongPollingBot;
import org.telegram.telegrambots.meta.api.methods.send.SendDocument;
import org.telegram.telegrambots.meta.api.methods.send.SendMessage;
import org.telegram.telegrambots.meta.api.objects.Document;
import org.telegram.telegrambots.meta.api.objects.InputFile;
import org.telegram.telegrambots.meta.api.objects.Update;
import org.telegram.telegrambots.meta.exceptions.TelegramApiException;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;

@Service
public class BotService extends TelegramLongPollingBot {

    private static final String FILE_ID_REGEX = "^.*\\/file\\/([\\w-]+)\\/.*$";


    @Override
    public String getBotToken() {
        // return your bot's API token here
        return "6055790040:AAF3lO6RxwixWCRBnU68kCiB3Bnm0Nko3nE";
    }

    @Override
    public String getBotUsername() {
        // return your bot's username here
        return "MahiBot";
    }

    @Override
    public void onUpdateReceived(Update update) {
        // handle incoming messages here
        if (update.hasMessage() && update.getMessage().hasText()) {
            String messageText = update.getMessage().getText();
            String chatId = update.getMessage().getChatId().toString();
            // send a reply message to the chat
            System.out.println(messageText);
            sendTextMessage(chatId, "You said : " + messageText);
        }

        if(update.hasMessage() && update.getMessage().hasDocument()){
            String chatId = update.getMessage().getChatId().toString();
            Document document = update.getMessage().getDocument();
            if(document.getFileName().endsWith("xlsx") ||document.getFileName().endsWith("XLSX")){
                String fileId = document.getFileId();
                String fileName = document.getFileName();
                sendTextMessage(chatId, "ትንሽ ይጠብቁ!");
                if (fileId != null) {
                    InputStream fileStream = downloadExcelFile(fileId);
                    if (fileStream != null) {
                        try {
                            File file = new File("samp.xlsx");
                            FileOutputStream outputStream = new FileOutputStream(file);
                            byte[] buffer = new byte[4096];
                            int bytesRead = 0;
                            while ((bytesRead = fileStream.read(buffer)) != -1) {
                                outputStream.write(buffer, 0, bytesRead);
                            }
                            outputStream.close();
                            fileStream.close();
                            ExcelMerger excelMerger = new ExcelMerger();
                            excelMerger.generateGpx("samp.xlsx", fileName);
                            String outputFileName = fileName.substring(0, fileName.length() - 4) + "gpx";
                            SendDocument document1 = new SendDocument(chatId, new InputFile(new File(outputFileName)));
                            execute(document1);
                        } catch (Exception e) {
                            e.printStackTrace();
                            sendTextMessage(chatId, "Error processing file");
                        }
                    } else {
                        sendTextMessage(chatId, "Error downloading file");
                    }
                }
            }
        }
    }

    public void sendTextMessage(String chatId, String text) {
        // send a text message to the specified chat
        // you can use the Telegram Bot API to send messages
        SendMessage sendMessage = new SendMessage(chatId, text);
        try{
            execute(sendMessage);
        } catch (TelegramApiException e){
            e.printStackTrace();
        }
    }
    private void sendFile(String chatId, String fileName) {
        try {
            InputFile file = new InputFile(fileName);
            SendDocument sendDocument = new SendDocument();
            sendDocument.setChatId(chatId);
            sendDocument.setDocument(file);
            execute(sendDocument);
        } catch (Exception e) {
            e.printStackTrace();
            sendTextMessage(chatId, "Error sending file");
        }
    }

    private String extractFileId(String messageText) {
        String[] parts = messageText.split(" ");
        for (String part : parts) {
            if (part.matches(FILE_ID_REGEX)) {
                return part.replaceAll(FILE_ID_REGEX, "$1");
            }
        }
        return null;
    }

    private InputStream downloadExcelFile(String fileId) {
        try {
            String fileUrl = getFileUrl(fileId);
            System.out.println(fileUrl);
            String accessToken = getBotToken();
            URL url = new URL("https://api.telegram.org/file/bot" + accessToken + "/" + fileUrl);
            System.out.println(url);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            int responseCode = connection.getResponseCode();
            if (responseCode == HttpURLConnection.HTTP_OK) {
                return connection.getInputStream();
            } else {
                System.out.println("Failed to download file. Response code: " + responseCode);
                return null;
            }
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    private String getFileUrl(String fileId) throws IOException {
        String apiUrl = "https://api.telegram.org/bot" + getBotToken() + "/getFile?file_id=" + fileId;
        URL url = new URL(apiUrl);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");
        int responseCode = connection.getResponseCode();
        if (responseCode == HttpURLConnection.HTTP_OK) {
            BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
            String response = reader.readLine();
            System.out.println(response);
//            System.out.println(JsonParser.parseString(response));
//            JsonObject json = JsonParser.parseString(response).getAsJsonObject();
//            System.out.println(json.getAsJsonPrimitive());
//            return json.getAsJsonPrimitive("result").getAsJsonObject().getAsJsonPrimitive("file_path").getAsString();
            ObjectMapper mapper = new ObjectMapper();
            JsonNode rootNode = mapper.readTree(response);
            JsonNode resultNode = rootNode.get("result");
            return resultNode.get("file_path").asText();
        } else {
            System.out.println("Failed to get file URL. Response code: " + responseCode);
            return null;
        }
    }

    private String getFileName(String fileId) throws IOException {
        String apiUrl = "https://api.telegram.org/bot" + getBotToken() + "/getFile?file_id=" + fileId;
        URL url = new URL(apiUrl);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");
        int responseCode = connection.getResponseCode();
        if (responseCode == HttpURLConnection.HTTP_OK) {
            BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
            String response = reader.readLine();
            JsonObject json = JsonParser.parseString(response).getAsJsonObject();
//            String filePath = json.getAsJsonPrimitive("result").getAsJsonObject().getAsJsonPrimitive("file_path").getAsString();
            ObjectMapper mapper = new ObjectMapper();
            JsonNode rootNode = mapper.readTree(response);
            JsonNode resultNode = rootNode.get("result");
            String filePath = resultNode.get("file_path").asText();
            return new File(filePath).getName();
        } else {
            System.out.println("Failed to get file name. Response code: " + responseCode);
            return null;
        }
    }

//    private InputStream downloadFileContent(String fileId) throws TelegramApiException {
//        org.telegram.telegrambots.meta.api.objects.File file = execute(new org.telegram.telegrambots.meta.api.methods.GetFile(fileId));
//        try {
//            return new ByteArrayInputStream(file.getContent);
//        } catch (IOException e) {
//            throw new TelegramApiException("Failed to download file content", e);
//        }
//    }
}
