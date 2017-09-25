<?php
/*
  Plugin Name: Regenerate post permalinks
  Plugin URI: https://wordpress.org/plugins/regenerate-post-permalinks/
  Description: This plugin can help you to regenerate all your permalinks based on the post titles.
  Author: Sandor Kovacs
  Version: 1.0.3
  Author URI: http://sandorkovacs.ro/en/
 */


add_action('admin_menu', 'register_simple_rpp_submenu_page');

function register_simple_rpp_submenu_page() {
    add_submenu_page(
            'options-general.php', __('Permalink regeneration'), __('Permalink regeneration'), 'manage_options', 'regenerate-post-permalink', 'rpp_callback');
}

/** Remove diacritics **/
function regenerate_post_clear_diacritics($str) {
    $table = array(
        'À' => 'A', 'Á' => 'A', 'Â' => 'A', 'Ã' => 'A', 'Ä' => 'A', 'Å' => 'A', 
        'Æ' => 'A', 'à' => 'a', 'á' => 'a', 'â' => 'a', 'ã' => 'a', 'ä' => 'a', 'å' => 'a', 'æ' => 'ae',
        'Č' => 'C', 'č' => 'c', 'Ć' => 'C', 'ć' => 'c', 'Ç' => 'C', 'ç' => 'c',
        'Ď' => 'D', 'ď' => 'd',
        'È' => 'E', 'É' => 'E', 'Ê' => 'E', 'Ë' => 'E', 'Ě' => 'E', 'è' => 'e', 
        'é' => 'e', 'ê' => 'e', 'ë' => 'e', 'ě' => 'e',
        'Ì' => 'I', 'Í' => 'I', 'Î' => 'I', 'Ï' => 'I', 'ì' => 'i', 'í' => 'i', 'î' => 'i', 'ï' => 'i',
        'Ñ' => 'N', 'ñ' => 'n',
        'Ò' => 'O', 'Ó' => 'O', 'Ô' => 'O', 'Õ' => 'O', 'Ö' => 'O', 'Ø' => 'O', 
        'ð' => 'o', 'ò' => 'o', 'ó' => 'o', 'ô' => 'o', 'õ' => 'o', 'ö' => 'o', 'ø' => 'o',
        'Ŕ' => 'R', 'Ř' => 'R', 'Ŕ' => 'R', 'ŕ' => 'r', 'ř' => 'r',
        'Š' => 'S', 'š' => 's', 'Ś' => 'S', 'ś' => 's',
        'Ť' => 'T', 'ť' => 't',
        'Ù' => 'U', 'Ú' => 'U', 'Û' => 'U', 'Ü' => 'U', 'ù' => 'u', 'ú' => 'u', 
        'û' => 'u', 'ü' => 'u',
        'Ý' => 'Y', 'ÿ' => 'y', 'ý' => 'y', 'ý' => 'y',
        'Ž' => 'Z', 'ž' => 'z', 'Ź' => 'Z', 'ź' => 'z',
        'Đ' => 'Dj', 'đ' => 'dj', 'Þ' => 'B', 'ß' => 's', 'þ' => 'b',
    );
    
    return  strtr($str, $table);
}

// Regenerate post permalink
function regenerate_post_permalink($post_type = 'post') {
    global $wpdb;

    $myrows = $wpdb->get_results("SELECT id, post_title FROM $wpdb->posts WHERE post_status = 'publish' AND post_type='$post_type' ");
    $counter = 0;
    foreach ($myrows as $pid) :
        $post_title = regenerate_post_clear_diacritics($pid->post_title);
        $guid = home_url() . '/' . sanitize_title_with_dashes($post_title);
        $sql = "UPDATE $wpdb->posts 
                     SET post_name = '" . sanitize_title_with_dashes($post_title) . "',
                         guid = '" . $guid . "'
               WHERE ID = $pid->id";
        $wpdb->query($sql);
        $counter++;
    endforeach;

    return $counter;
}

/** POST META MANAGER PLUGIN PAGE * */
function rpp_callback() {
    ?>
    <div class="wrap" id='simple-sf'>
        <div class="icon32" id="icon-options-general"><br></div><h2><?php _e('Permalinks regeneration'); ?></h2>
    <?php _e('<p>Simply select the post type and click on the Regenerate permalinks button.It will regenerate all the permalinks based on title.  </p>
        <p>Eg. <strong>"This is my article title"</strong> will have the following permalink:  <em>"this-is-my-article-title"</em></p>') ?>

        <?php
        if (isset($_POST['submit-regenerate-permalinks'])) :
            $counter = regenerate_post_permalink($_POST['rpp-post-type']);
            ?>

            <p>  <?php printf(__('%1$s permalinks have been regenerated for all posts having type: <strong>%2$s</strong>'), $counter, $_POST['rpp-post-type']); ?></p>

        <?php else:
            ?>

            <br/>
            <form action="" method="post" name="rpp">

                <table class="form-table">

                    <tr>
                        <th scope="row"><strong><?php _e('Select post type'); ?></strong></th>
                        <td>      
                            <select name="rpp-post-type">
        <?php
        $post_types = get_post_types(array('public' => true), 'names');

        foreach ($post_types as $post_type) {

            echo '<option value="' . $post_type . '">' . $post_type . '</option>';
        }
        ?>                
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <th scope="row">&nbsp;</th>
                        <td>      
                            <input type='submit' 
                                   onclick='if (!window.confirm("<?php _e('Are you really sure?') ?>"))
                                                       return false'
                                   name='submit-regenerate-permalinks' value='<?php _e('Regenerate permalinks') ?>' />
                        </td>
                    </tr>

                </table>

            </form>

    <?php endif; ?>
    </div>

        <?php
    }
    