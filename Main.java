package exelReader;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Optional;
import java.util.stream.Stream;

import exelReader.reader.ReaderXlsx;

public class Main {
	private boolean work = true;
	private boolean iscopy = true;
	
	public void work() {
		
		Path path = Paths.get("C:\\Users\\Tomek\\Documents\\arkusze");
		Path pathFile = null;
		while(work) {
			if(iscopy) {
				try (Stream<Path> paths = Files.list(path)){
					Optional<Path> opath = paths.filter(p -> p.toString().endsWith("xlsx")).findAny();
					if(opath.isPresent()) {
						pathFile = opath.get();
						File xlsxFile = new File(pathFile.toString());
						Thread thread = new Thread(new ReaderXlsx(xlsxFile,this));
						thread.start();
						iscopy = false;
					}else {
						System.err.println("No xlsx file in directory");
					}
					
				}catch(IOException e) {
					System.err.println("No Such directory");
				}
			}
			
		}
	}
	
	public synchronized void stopWork() {
		this.work = false;
	}
	public  synchronized void setCopy() {
		this.iscopy = true;
	}

	public static void main(String[] args) {
		System.out.println("xlnx parser start");
		Main main = new Main();
		main.work();
	}
}
