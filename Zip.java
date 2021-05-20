package swagger2word.utils;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveInputStream;
import org.apache.commons.compress.archivers.zip.ZipArchiveOutputStream;
import org.apache.commons.compress.utils.IOUtils;

// https://itsallbinary.com/apache-commons-compress-simplest-zip-zip-with-directory-compression-level-unzip/
public final class Zip {

  public static final int ZIP_FLAT_MODE = 1;
  public static final int ZIP_RELATIVE_MODE = 2;

  public static void compress(File source, File zip, int mode) throws IOException {
    // Create zip file stream.
    try (ZipArchiveOutputStream archive = new ZipArchiveOutputStream(new FileOutputStream(zip))) {
      // Walk through files, folders & sub-folders.
      addArchiveEntry(archive, source, mode);
      // Complete archive entry addition.
      archive.finish();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  public static void compressMultiDir(List<File> sources, File zip, int mode) throws IOException {
    // Create zip file stream.
    //    if (zip.exists()) {
    //      FileUtils.forceDelete(zip);
    //    }
    try (ZipArchiveOutputStream archive = new ZipArchiveOutputStream(new FileOutputStream(zip))) {
      for (File source : sources) {
        // Walk through files, folders & sub-folders.
        addArchiveEntry(archive, source, mode);
      }
      // Complete archive entry addition.
      archive.finish();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  private static void addArchiveEntry(ZipArchiveOutputStream archive, File source, int mode)
      throws IOException {

    Files.walk(source.toPath())
        .forEach(
            p -> {
              File file = p.toFile();
              // Directory is not streamed, but its files are streamed into zip file with folder
              // in it's path
              if (!file.isDirectory()) {
                // Add file to zip archive
                ZipArchiveEntry entry_1 =
                    new ZipArchiveEntry(
                        file,
                        mode == ZIP_FLAT_MODE
                            ? getFlatZipEntryName(file)
                            : getRelativeZipEntryName(file));
                try (FileInputStream fis = new FileInputStream(file)) {
                  archive.putArchiveEntry(entry_1);
                  IOUtils.copy(fis, archive);
                  archive.closeArchiveEntry();
                } catch (IOException e) {
                  e.printStackTrace();
                }
              }
            });
  }

  public static void extract() {
    // Create zip file stream.
    try (ZipArchiveInputStream archive =
        new ZipArchiveInputStream(
            new BufferedInputStream(new FileInputStream("output/sample-fast.zip")))) {

      ZipArchiveEntry entry;
      while ((entry = archive.getNextZipEntry()) != null) {
        // Print values from entry.
        // System.out.println(entry.getName());
        // System.out.println(entry.getMethod()); // ZipEntry.DEFLATED is int 8

        File file = new File("output/" + entry.getName());
        // System.out.println("Unzipping - " + file);
        // Create directory before streaming files.
        String dir =
            file.toPath().toString().substring(0, file.toPath().toString().lastIndexOf("\\"));
        Files.createDirectories(new File(dir).toPath());
        // Stream file content
        IOUtils.copy(archive, new FileOutputStream(file));
      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  private static String getRelativeZipEntryName(File file) {
    return Paths.get(file.getParentFile().getName(), file.getName()).toString();
  }

  private static String getFlatZipEntryName(File file) {
    return Paths.get(file.getName()).toString();
  }
}
