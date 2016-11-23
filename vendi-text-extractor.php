<?php
/*
Plugin Name: Vendi Text Extractor
Version: 0.0.2
Author: Vendi Advertising (Chris Haas)
*/

add_action( 'add_attachment', 'vendi__add_attachment' );

define( 'VENDI_TOOL_PATH_PDF_TO_TEXT',   dirname( __FILE__ ) . '/bin/pdftotext/pdftotext' );
define( 'VENDI_TOOL_PATH_DOC_TO_TEXT',   dirname( __FILE__ ) . '/bin/catdoc/catdoc' );
define( 'VENDI_TOOL_PATH_DOCX_TO_TEXT',  dirname( __FILE__ ) . '/bin/docx2txt/docx2txt.pl' );

define( 'VENDI_IDEAL_FILE_PERMS',       0664 );

define( 'VENDI_DEBUG_TEXT_EXTRACTOR', false );

function vendi__add_attachment( $post_ID )
{
    //Sanity check for a value, just in case
    if ( ! $post_ID )
    {
        return;
    }

    //Get the actual post object and bail if there isn't one, just in case
    $post = get_post( $post_ID );
    if ( ! $post )
    {
        return;
    }

    //Make sure we have a mime type and bail if we don't
    if ( ! property_exists( $post, 'post_mime_type' ) || ! isset( $post->post_mime_type ) )
    {
        return;
    }

    //Get WordPress's upload directory so that we can create and store a temporary file
    $upload_dir = wp_upload_dir();

    //Get the uploaded file's path
    $source_file_path = get_attached_file( $post_ID, true );

    //Create a unique temporary file to extract text to
    $destination_file_path = tempnam( $upload_dir['path'], 'VENDI_' );

    //TODO: We should handle this error
    if ( ! $destination_file_path )
    {
        return;
    }

    $cmd = false;

    //Attempt to set proper permissions on the temporary file
    chmod( $destination_file_path, VENDI_IDEAL_FILE_PERMS );

    switch ( $post->post_mime_type )
    {
        //PDF
        case 'application/pdf':
            $cmd = escapeshellcmd ( sprintf( '%1$s -htmlmeta -eol dos %2$s %3$s', VENDI_TOOL_PATH_PDF_TO_TEXT, $source_file_path, $destination_file_path ) );

            break;

        //DOC
        case 'application/msword':
            $cmd_left  = escapeshellcmd( sprintf( '%1$s -dutf-8 %2$s', VENDI_TOOL_PATH_DOC_TO_TEXT, $source_file_path ) );
            $cmd_right = escapeshellcmd( $destination_file_path );
            $cmd = $cmd_left . ' > ' . $cmd_right;

            break;

        //DOCX
        case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            $cmd_left    = escapeshellcmd( VENDI_TOOL_PATH_DOCX_TO_TEXT );
            $cmd_middle  = escapeshellcmd( $source_file_path );
            $cmd_right   = escapeshellcmd( $destination_file_path );
            $cmd = $cmd_left . ' < ' . $cmd_middle . ' > ' . $cmd_right;

            break;

        //PPTX
        //http://superuser.com/a/706620/24444
        case 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
            $cmd = sprintf( 'unzip -qc "%1$s" ppt/slides/slide*.xml | grep -oP \'(?<=\<a:t\>).*?(?=\</a:t\>)\' > %2$s', escapeshellcmd( $source_file_path ), escapeshellcmd( $destination_file_path ) );

            break;

        //XLS and XLSX
        case 'application/vnd.ms-excel':
        case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            require_once( 'bin/phpexcel/Classes/PHPExcel.php' );

            $objPHPExcel = PHPExcel_IOFactory::load( $source_file_path );
            $objWriter = PHPExcel_IOFactory::createWriter( $objPHPExcel, 'HTML' );
            $objWriter->writeAllSheets();

            $file = fopen( $destination_file_path, 'w' );
            fwrite( $file, $objWriter->generateSheetData() );
            fclose( $file );

            unset( $objWriter );
            unset( $objPHPExcel );

            break;

        default:
            //echo $post->post_mime_type;
            return;
    }

    $debug_bugger = array();

    if( false !== $cmd )
    {
        vendi__debug_text_extractor( $cmd );

        //Output will be all text outputted by the function, useful for debugging
        $output = array();

        //This will be the error code returned from the command
        $return_val = null;

        //Call our command
        exec( $cmd, $output, $return_val );

        vendi__debug_text_extractor( $output );
        vendi__debug_text_extractor( $return_val );

        //TODO: Better error handling
        //Zero is success, anything else is failure
        if( 0 !== $return_val )
        {
            echo sprintf( 'exec command returned error code %1$s', $return_val );
            return;
        }
    }

    //Get the content that was exported
    $extracted_text = file_get_contents( $destination_file_path );
    if( ! $extracted_text )
    {
        echo 'There was an unknown error reading from the text extract file';
        return;
    }

    vendi__debug_text_extractor( $extracted_text );

    $new_post = array(
                        'ID'            => $post_ID,
                        'post_content'  => $extracted_text
                    );

    wp_update_post( $new_post );

    //Attempt to cleanup our file
    unlink( $destination_file_path );
}

function vendi__debug_text_extractor( $obj )
{
    if( defined( 'VENDI_DEBUG_TEXT_EXTRACTOR' ) && true === VENDI_DEBUG_TEXT_EXTRACTOR )
    {
        echo '<pre>';
        print_r( $obj );
        echo '</pre>';
        echo '<hr />';
    }
}