function get_attr_as_boolean(
    p_region in apex_plugin.t_region,
    p_attribute in varchar2
)
return boolean
is
    l_attribute varchar2(10);
begin
    l_attribute := p_region.attributes.get_varchar2(p_attribute);
    return (l_attribute is not null and l_attribute = 'Y');
end;

procedure render (
    p_plugin in            apex_plugin.t_plugin,
    p_region in            apex_plugin.t_region,
    p_param  in            apex_plugin.t_region_render_param,
    p_result in out nocopy apex_plugin.t_region_render_result )
is 
    l_region_id             varchar2(50);  
begin
    if apex_application.g_debug then
        apex_plugin_util.debug_region(p_plugin => p_plugin, p_region => p_region);
    end if;
    l_region_id := apex_escape.html_attribute(p_region.static_id);

    sys.htp.p('<div id="' || l_region_id || '_sv" class="lib4x-SV"></div>');
 
    -- When specifying the library declaratively, it fails to load the minified version. So using the API:
    apex_javascript.add_library(
          p_name      => 'ig-spreadsheetview',
          p_check_to_add_minified => true,
          --p_directory => '#WORKSPACE_FILES#javascript/',          
          p_directory => p_plugin.file_prefix || 'js/',
          p_version   => NULL
    );  

    apex_css.add_file (
        p_name => 'jspreadsheet.themes',
        --p_directory => '#WORKSPACE_FILES#css/themes/'
        p_directory => p_plugin.file_prefix || 'css/' 
    );

    apex_css.add_file (
        p_name => 'ig-spreadsheetview',
        --p_directory => '#WORKSPACE_FILES#css/'
        p_directory => p_plugin.file_prefix || 'css/' 
    );    

    apex_javascript.add_onload_code(
        p_code => apex_string.format(
            'lib4x.axt.ig.spreadsheetView._init("%s", "%s", "%s", "%s", "%s", '
            , l_region_id
            , apex_escape.html(p_region.attributes.get_varchar2('attr_enable_on'))  
            , apex_escape.html(p_region.attributes.get_varchar2('attr_ig_static_id'))  
            , apex_escape.html(p_region.attributes.get_varchar2('attr_filter_class'))       
            , apex_escape.html(p_region.attributes.get_varchar2('attr_pagination_type'))     
        ) || p_region.init_javascript_code || ');'
    );    
end;
