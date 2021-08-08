package com.example.lucifer.mypoidemo;


import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
//import android.support.annotation.NonNull;
//import android.support.annotation.Nullable;
//import android.support.v4.app.Fragment;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.fragment.app.Fragment;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import java.io.File;
import java.io.IOException;


/**
 * A simple {@link Fragment} subclass.
 */
public class fileChooserFragment extends Fragment implements View.OnClickListener {

    EditText txtFileName;
    Button btnChooser;
    TextView lbFilePath;

    private static final int MY_RESULT_CODE_FILECHOOSER = 2000;

    public fileChooserFragment() {
        // Required empty public constructor
    }


    @Override
    public View onCreateView(LayoutInflater inflater, ViewGroup container,
                             Bundle savedInstanceState) {
        // Inflate the layout for this fragment
        return inflater.inflate(R.layout.file_chooser_fragment, container, false);
    }

    @Override
    public void onViewCreated(@NonNull View view, @Nullable Bundle savedInstanceState) {
        super.onViewCreated(view, savedInstanceState);
        txtFileName=(EditText)view.findViewById(R.id.edittextFileName );
        lbFilePath = (TextView)view.findViewById(R.id.textviewFilePath);

        lbFilePath.setVisibility(View.INVISIBLE);

        btnChooser = (Button)view.findViewById(R.id.buttonChooser);
        btnChooser.setOnClickListener(this);
    }

    @Override
    public void onClick(View view) {
        Intent intentOpenFile = new Intent();
        intentOpenFile.setAction(Intent.ACTION_GET_CONTENT);
        intentOpenFile.setType("*/*");
        intentOpenFile.addCategory(Intent.CATEGORY_OPENABLE);

        Intent chooser = Intent.createChooser(intentOpenFile,"Chọn file để mở");

        if (intentOpenFile.resolveActivity(getActivity().getPackageManager())!=null){
            startActivityForResult(chooser,MY_RESULT_CODE_FILECHOOSER);
        }
    }

    @Override
    public void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        Uri uriReturn = data.getData();
        String filePath = null;
        String fileName = null;

        File file = null;

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M ){
            Cursor c = null;
            String[] projection={MediaStore.Files.FileColumns.DISPLAY_NAME};
            //c = getActivity().getContentResolver().query(MediaStore.Files.getContentUri("external"),projection,null,null,null);
            c = getActivity().getContentResolver().query(uriReturn, projection,null,null,null);
            if(c!=null){
                c.moveToFirst();
                fileName = c.getString(c.getColumnIndexOrThrow(MediaStore.Files.FileColumns.DISPLAY_NAME));
                if(Environment.getExternalStorageState().equalsIgnoreCase("mounted")){
                    filePath = Environment.getExternalStorageDirectory().toString();
                    file = new File(filePath + "/Download/"+fileName);
                }
            }
        }else{
            file = new File(uriReturn.getPath());
        }

        if (file!=null){
            fileName = file.getName();
            filePath = file.getPath();

            lbFilePath.setText(filePath);
            txtFileName.setText(fileName);

        }else{
            Toast toast = Toast.makeText(getActivity(),"Chưa lấy được file",Toast.LENGTH_SHORT);
            toast.show();
        }

    }

}
