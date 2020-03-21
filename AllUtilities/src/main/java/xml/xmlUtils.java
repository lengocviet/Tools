package xml;

import java.io.File;
import java.io.IOException;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Node;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;

public class xmlUtils {

    public  static void readxml(String path) throws ParserConfigurationException, IOException, SAXException {

        File fxMFile = new File(path);
        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder ddBuilder = dbFactory.newDocumentBuilder();
        Document doc = ddBuilder.parse(fxMFile);
        doc.getDocumentElement().normalize();
        String rootNode = doc.getDocumentElement().getNodeName();
        NodeList nList = doc.getElementsByTagName("Sheet");
        for (int i = 0;  i< nList.getLength();i++) {

            Node nNode = nList.item(i);
            //System.out.print(nNode.getNodeName());
            if (nNode.getNodeType() ==Node.ELEMENT_NODE)
            {
                NodeList nChildNodes = nNode.getChildNodes();
                for (int j =0; j< nChildNodes.getLength(); j++) {
                    Node nodeItems = nChildNodes.item(j);
                    if (nodeItems.getNodeType()==Node.ELEMENT_NODE) {
                        Element eElement = (Element) nodeItems;

                        System.out.println("Name : " + eElement.getAttribute("name"));
                        System.out.println("firstCol : " + eElement.getAttribute("firstCol"));
                        System.out.println("endCol : " + eElement.getAttribute("endCol"));

                    }
                }
            }

        }

    }
}
