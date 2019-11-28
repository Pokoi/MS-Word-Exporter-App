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

        documentTitleContainer   = (EditText)   findViewById(R.id.editTextTitle);
        documentContentContainer = (EditText)   findViewById(R.id.editTextContent);
        generateDocumentButton   = (Button)     findViewById(R.id.button);

        generateDocumentButton.setOnClickListener(new View.OnClickListener()
        {
            @Override
            public void onClick(View v)
            {
                MSWordJavaObject document = new MSWordJavaObject(documentTitleContainer.getText().toString());
                document.AddParagraph(documentContentContainer.getText().toString());
                try
                {
                    File file = new File(getExternalFilesDir(null), "doc.docx");
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
