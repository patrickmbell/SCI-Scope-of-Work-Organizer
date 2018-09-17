import java.io.*;
import java.lang.*;
import java.util.*;

public class BufferedInput {

    public ListItems[] x;


    class ListItems{
        String active = null;
        String serviceType = null;
        String description = null;

    }

    public BufferedInput() {

    }

    ListItems[] scanDoc() throws IOException, FileNotFoundException
    {
            x = new ListItems[363];
            for(int i = 0; i < x.length; i++)
            {
                x[i] = new ListItems();
            }
            int i = 0;
            BufferedReader reader = new BufferedReader(new FileReader("Ordered.txt"));
            String line = " ";
            while((line = reader.readLine()) != null)
            {
                if(line.contains("Status"))
                {
                    line = reader.readLine();
                    i++;
                    continue;
                }
                else
                {
                    if(line.contains("Active")) {
                        x[i].active = "Active";
                    }
                    else if(line.contains("Not-active"))
                    {
                        x[i].active = "Not-Active";
                    }

                    else if(line.contains("Service\tAnnual Maintenance"))
                    {
                        x[i].serviceType = "Annual Maintenance";
                    }

                    else if(line.contains("BROCHURES"))
                    {
                        x[i].serviceType = "BROCHURES";
                    }

                    else if(line.contains("Building Cost"))
                    {
                        x[i].serviceType = "Building Cost";
                    }
                    else if(line.contains("Building Sale"))
                    {
                        x[i].serviceType="Building Sale";
                    }
                    else if(line.contains("Cabinets & Vanities"))
                    {
                        x[i].serviceType="Cabinets & Vanities";
                    }
                    else if(line.contains("Clean Up"))
                    {
                        x[i].serviceType = "Clean Up";
                    }

                }

                i++;
            }


            return x;
    }

    public ListItems[] getList()
    {
        return x;
    }

    public String getActive(int index)
    {
        return x[index].active;
    }


}
