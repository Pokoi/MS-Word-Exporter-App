package com.pokoidev.mswordexporter;

import androidx.appcompat.app.AppCompatActivity;

import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import java.io.File;

public class MainActivity extends AppCompatActivity {

    EditText documentTitleContainer;
    EditText documentContentContainer;
    Button   generateDocumentButton;

    @Override
    protected void onCreate(Bundle savedInstanceState) {

        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        documentContentContainer = (EditText)   findViewById(R.id.editTextContent);
        documentTitleContainer   = (EditText)   findViewById(R.id.editTextTitle);
        generateDocumentButton   = (Button)     findViewById(R.id.button);

        generateDocumentButton.setOnClickListener(new View.OnClickListener()
        {
            @Override
            public void onClick(View v)
            {
                /*
                MSWordMetaData metaData = new MSWordMetaData(
                                                            "Informe",
                                                            "Informe de accesibilidad",
                                                            "accesibilidad",
                                                            "Evaluaciones del técnico especializado",
                                                            " ",
                                                            "Insertar nombre del técnico",
                                                            " ",
                                                            " ",
                                                            "Fundación Once"
                                                            );

                MSWordWritableDocument document = new MSWordWritableDocument(metaData);
                document.AddHeading(documentContentContainer.getText().toString(), 1);
                document.AddParagraph(documentContentContainer.getText().toString());
                try
                {
                    File file = new File (getExternalFilesDir(null), document.getMetaData().getTitle() + ".doc");
                    document.ExportFile(file);
                    Toast.makeText(getApplicationContext(),"OK", Toast.LENGTH_LONG).show();
                } catch (Exception e)
                {
                    Toast.makeText(getApplicationContext(),"NOT OK", Toast.LENGTH_LONG).show();
                    e.printStackTrace();
                }*/

                MSWordTemplate document = new MSWordTemplate(getExternalFilesDir(null) +"/"+ "template.xml");
                document.Replace("centerName", "Holiwi");

                try
                {
                    File file = new File (getExternalFilesDir(null),  "informe.doc");
                    document.ExportFile(file);
                    Toast.makeText(getApplicationContext(),"OK", Toast.LENGTH_LONG).show();
                } catch (Exception e)
                {
                    Toast.makeText(getApplicationContext(),"NOT OK", Toast.LENGTH_LONG).show();
                    e.printStackTrace();
                }


            }
        });
    }
}
