require 'ews_service'

@ews_service = EwsService.new
@calendar_xml = @ews_service.get_calendar_xml
puts @ews_service.parse_calendar_for_meetings(@calendar_xml)
