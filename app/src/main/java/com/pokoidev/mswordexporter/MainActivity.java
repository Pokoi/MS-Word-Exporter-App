package com.pokoidev.mswordexporter;

import androidx.appcompat.app.AppCompatActivity;

import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;

import java.io.File;

public class MainActivity extends AppCompatActivity {

    EditText documentTitleContainer;
    EditText documentContentContainer;
    Button generateDocumentButton;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        documentTitleContainer   = (EditText) findViewById(R.id.editTextTitle);
        documentContentContainer = (EditText) findViewById(R.id.editTextContent);
        generateDocumentButton   = (Button) findViewById(R.id.button);

        generateDocumentButton.setOnClickListener(new View.OnClickListener()
        {
            @Override
            public void onClick(View v)
            {
                MSWordJavaObject document = new MSWordJavaObject(documentTitleContainer.getText().toString());
                document.AddParagraph(documentContentContainer.getText().toString());
                try {
                    File file = new File(getExternalFilesDir(null), "doc.docx");
                    document.ExportFile(file);
                } catch (Exception e) {
                    e.printStackTrace();
                }

            }
        });
    }
}
