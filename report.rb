# encoding: utf-8

require 'getoptlong'
require 'date'
require 'time'
require 'pathname'

$super_verbose = false
$verbose       = false

class Event < Struct.new(:listid, :name, :title, :length, :live, :hershim, :simanim, :long_filename, :saturday, :illegal, :too_short, :morning_lesson_hershim)
  PRIME_TIME   = ('19:00:00' .. '23:00:00')
  ILLEGAL_TIME = ('01:00:00' .. '06:00:00')

  def print_row(time)
    self.too_short = self.length < 10 && long_filename !~ /_program_|_lesson_/
    self.saturday  = time.wday == 6

    length_sec = time + self.length

    $stderr.puts "#{time} #{name} #{title} #{self.length}" if $super_verbose

    if PRIME_TIME.cover?(time.strftime('%H:%M:%S'))
      if PRIME_TIME.cover?((time + self.length).strftime('%H:%M:%S'))
        $stderr.puts 'starts inside, ends inside PRIME_TIME' if $super_verbose
        print_one_line time, self.length, true, false
      else
        $stderr.puts 'starts inside, ends after PRIME_TIME' if $super_verbose
        finish        = Time.parse("#{time.to_date.to_s} 23:00:00")
        inside_length = finish - time
        print_one_line time, inside_length, true, false
        print_one_line finish, self.length - inside_length, false, false
      end
    else
      if PRIME_TIME.cover?(length_sec.strftime('%H:%M:%S'))
        $stderr.puts 'starts before, ends inside PRIME_TIME' if $super_verbose
        start         = Time.parse("#{time.to_date.to_s} 19:00")
        before_length = start - time

        print_one_line time, before_length, false, false
        print_one_line start, self.length - before_length, true, false
      else
        if ILLEGAL_TIME.cover?(time.strftime('%H:%M:%S'))
          if ILLEGAL_TIME.cover?(length_sec.strftime('%H:%M:%S'))
            $stderr.puts 'starts inside, ends inside ILLEGAL_TIME' if $super_verbose
            print_one_line time, self.length, false, true
          else
            $stderr.puts 'starts inside, ends after ILLEGAL_TIME' if $super_verbose
            finish        = Time.parse("#{time.to_date.to_s} 06:00")
            inside_length = finish - time
            print_one_line time, inside_length, false, true
            print_one_line finish, self.length - inside_length, false, false
          end
        else
          if ILLEGAL_TIME.cover?(length_sec.strftime('%H:%M:%S'))
            $stderr.puts 'starts before, ends inside ILLEGAL_TIME' if $super_verbose
            if time.to_date != length_sec.to_date
              $stderr.puts ' and crosses midnight!!!' if $super_verbose
              start = Time.parse("#{(time.to_date + 1).to_s} 01:00")
            else
              start = Time.parse("#{time.to_date.to_s} 01:00")
            end
            before_length = start - time

            print_one_line time, before_length, false, false
            print_one_line start, self.length - before_length, false, true
          else
            $stderr.puts 'starts outside, ends outside PRIME_TIME & ILLEGAL_TIME' if $super_verbose
            print_one_line time, self.length, false, false
          end
        end
      end
    end
  end

  def print_one_line(time, part_length, prime, illegal, length = self.length)
    id = case
           when illegal
             's65' # grey
           when (long_filename) =~ /(logo|promo|patiach|sagir|luz)/
             's67' # green
           when too_short
             's64' # red
           when saturday
             's62' # yellow
           when prime
             's63' # blue
           else
             's61' # none
         end
    puts <<-ROW
      <Row>
        <Cell ss:StyleID="#{id}"><Data ss:Type="String">#{listid}</Data></Cell>
        <Cell><Data ss:Type="String">#{time.to_date}</Data></Cell>
        <Cell><Data ss:Type="String">#{time.strftime '%H:%M:%S'}</Data></Cell>
        <Cell><Data ss:Type="String">#{(name.empty? || name == 'Not defined') ? title : name}</Data></Cell>
        <Cell><Data ss:Type="String">#{(title.empty? || title == 'Not defined') ? name : title}</Data></Cell>
        <Cell ss:StyleID="s66"><Data ss:Type="Number">#{'%02d.00' % ((part_length / 60).round)}</Data></Cell>
        <Cell ss:StyleID="s66"><Data ss:Type="Number">#{'%02d.00' % ((length / 60).round)}</Data></Cell>
        <Cell ss:StyleID="s66"><Data ss:Type="String">#{Time.at(part_length).utc.strftime('%H:%M:%S').to_s}</Data></Cell>
        <Cell ss:StyleID="s66"><Data ss:Type="String">#{Time.at(length).utc.strftime('%H:%M:%S').to_s}</Data></Cell>
        <Cell><Data ss:Type="String">#{prime ? 'כן' : 'לא'}</Data></Cell>
        <Cell><Data ss:Type="String">#{live ? 'כן' : 'לא'}</Data></Cell>
        <Cell><Data ss:Type="String">#{hershim ? 'כן' : 'לא'}</Data></Cell>
        <Cell><Data ss:Type="String">#{morning_lesson_hershim ? 'כן' : 'לא'}</Data></Cell>
        <Cell><Data ss:Type="String">#{simanim ? 'כן' : 'לא'}</Data></Cell>
        <Cell><Data ss:Type="String">#{long_filename}</Data></Cell>
      </Row>
    ROW
  end

  def self.print_header
    puts <<-HEADER
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:html="http://www.w3.org/TR/REC-html40">
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal">
     <Alignment ss:Vertical="Bottom"/>
     <Borders/>
     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
     <Interior/>
     <NumberFormat/>
     <Protection/>
    </Style>
    <Style ss:ID="s61">
    </Style>
    <Style ss:ID="s62">
     <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
    </Style>
    <Style ss:ID="s63">
     <Interior ss:Color="#0000FF" ss:Pattern="Solid"/>
    </Style>
    <Style ss:ID="s64">
     <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
    </Style>
    <Style ss:ID="s65">
     <Interior ss:Color="#666666" ss:Pattern="Solid"/>
    </Style>
    <Style ss:ID="s66">
     <NumberFormat ss:Format="Fixed"/>
    </Style>
    <Style ss:ID="s67">
     <Interior ss:Color="#00FF00" ss:Pattern="Solid"/>
    </Style>
  </Styles>
  <Worksheet ss:Name="Sheet1">
    <Table>
      <Row>
        <Cell><Data ss:Type="String">ID</Data></Cell>
        <Cell><Data ss:Type="String">תאריך שידור</Data></Cell>
        <Cell><Data ss:Type="String">שעת שידור</Data></Cell>
        <Cell><Data ss:Type="String">שם סדרה</Data></Cell>
        <Cell><Data ss:Type="String">שם פרק</Data></Cell>
        <Cell><Data ss:Type="String">אורך פרק (דק')</Data></Cell>
        <Cell><Data ss:Type="String">אורך מלא (דק')</Data></Cell>
        <Cell><Data ss:Type="String">אורך פרק</Data></Cell>
        <Cell><Data ss:Type="String">אורך מלא</Data></Cell>
        <Cell><Data ss:Type="String">פריים טיים</Data></Cell>
        <Cell><Data ss:Type="String">האם שידור חי או לא</Data></Cell>
        <Cell><Data ss:Type="String">האם היו כתוביות - כן/לא</Data></Cell>
        <Cell><Data ss:Type="String">כתוביות שיעור בוקר - כן/לא</Data></Cell>
        <Cell><Data ss:Type="String">האם תכנית עם שפת הסימנים - כן/לא</Data></Cell>
        <Cell><Data ss:Type="String">שם קובץ</Data></Cell>
      </Row>
    HEADER
  end

  def self.print_footer
    puts <<-FOOTER
    </Table>
  </Worksheet>
</Workbook> 
    FOOTER
  end
end

class Playlist
  MORNING_LESSON = ('02:45:00' .. '06:00:00')

  def initialize(start, finish, playlists_dir, hershim_dir)
    @playlisttc = '' # current time

    @start      = start
    @finish     = finish
    @date_range = @start..@finish

    @playlists_dir = playlists_dir
    @hershim_list  = Pathname.glob("#{hershim_dir}/*").map { |f| f.basename('.sub').to_s }

    start_file = start - 1
    file_range = start_file..finish

    @playlists = Dir["#{playlists_dir}/*"].sort_by { |name| name }.map do |playlist|
      name_parts = File.basename(playlist, '.ply')
      name_date  = Date.parse(name_parts.gsub('_', '-'))
      file_range.cover?(name_date) ? [playlist, name_date] : nil
    end.compact

    @event = Event.new
  end

  def parse
    Event.print_header unless $super_verbose

    @playlists.each do |playlist|
      parse_file(playlist[0], playlist[1])
    end

    Event.print_footer unless $super_verbose
  end

  private

  def parse_file(playlist, date)
    $stderr.puts "PLAYLIST #{playlist}" if $verbose
    $lineno        = 0 if $super_verbose
    performer      = 0
    live_performer = 0

    File.open(playlist, 'r') do |fd|
      while (line = fd.gets)
        if $super_verbose
          $lineno += 1
          $stderr.puts "Line ##{$lineno}"
        end
        line = line.force_encoding('Windows-1255').encode('UTF-8').chomp
        case line
          when /^#PLAYLISTTC/ # Once at top of file
            puts line if $super_verbose
            @playlisttc = Time.parse "#{date} #{line.split(' ').last}"
          when /^#LISTID (\d+)/
            puts line if $super_verbose
            @event.listid = $1
            @event.name   = ''
            @event.title  = ''
            performer     = 0
          when /^#METADATA Name\s+(.*)/
            puts line if $super_verbose
            @event.name = $1
          when /^#METADATA Title\s+(.*)/
            puts line if $super_verbose
            @event.title = $1
          when /^#PERFORMER\s+(.*)/
            puts line if $super_verbose
            begin
              performer = live_performer = Time.parse("#{@playlisttc.to_date.to_s} #{$1}")
            rescue
              performer = -1
            end
          when /^#EVENT WAITTO\s+(.*)/
            puts line if $super_verbose
            end_time = Time.parse("#{@playlisttc.to_date.to_s} 00:00") + $1.to_f
            unless @date_range.cover?(end_time.to_date)
              @playlisttc = end_time
              next
            end

            @event.name          = MORNING_LESSON.cover?(@playlisttc.strftime('%H:%M:00')) ? 'שיעור בוקר' : 'שידור חי'
            @event.title         = 'שידור חי'
            @event.live          = true
            @event.hershim       = false
            @event.simanim       = false
            @event.long_filename = ''
            @event.length        = end_time.to_time - @playlisttc.to_time
            @event.print_row @playlisttc

            @playlisttc = end_time
          when /^#EVENT WAIT\s+(.*)/, /^#LIVE_STREAM LIVE.+;(.*)/
            puts line if $super_verbose
            end_time = live_performer + $1.to_f

            @event.name          = MORNING_LESSON.cover?(@playlisttc.strftime('%H:%M:00')) ? 'שיעור בוקר' : 'שידור חי'
            @event.title         = 'שידור חי'
            @event.live          = true
            @event.hershim       = false
            @event.simanim       = false
            @event.long_filename = ''
            @event.length        = end_time.to_time - @playlisttc.to_time
            @event.print_row @playlisttc

            @playlisttc = end_time
          when /^"/ # Last line of block
            puts line if $super_verbose
            parts = line.split('; ')

            long              = File.basename(parts[0].split('"')[1].split('\\').last, '.mpg')
            play_time_trimmed = parts[1].to_f
            play_time         = parts[2].to_f - play_time_trimmed
            short             = parts[4]

            # Event happened out of requested dates
            unless @date_range.cover?(@playlisttc.to_date)
              $stderr.puts "#{@event.listid} SKIP #{@playlisttc} [time] [performer=#{performer}] #{short}" if $verbose
              @playlisttc += play_time
              next
            end

            # if long =~ /(logo|promo|patiach|sagir|luz)/
            #   $stderr.puts "#{@event.listid} SKIP #{@playlisttc} [#{$1}] [performer=#{performer}] #{short}" if $verbose
            #   @playlisttc += play_time
            #   next
            # end

            @event.simanim = long =~ /_hrsh$|_hrsh-[\da-f]+$/
            $stderr.puts "#{@event.listid} SIMANIM #{@playlisttc} #{short} #{long}" if $verbose && @event.simanim

            find_hershim(long)
            $stderr.puts "#{@event.listid} HERSHIM #{@playlisttc} #{short} #{long}" if $verbose && @event.hershim

            long =~ /(\d\d\d\d-\d\d-\d\d)/
            file_date = Date.parse($1) rescue Date.new

            @event.live          = (Date.parse((@playlisttc - 1).to_date.to_s) .. Date.parse((@playlisttc + 1).to_date.to_s)).cover?(file_date)
            @event.name          = @event.name.empty? ? short : @event.name
            @event.length        = play_time
            @event.long_filename = long

            @event.print_row @playlisttc

            @playlisttc += play_time
          else
        end
      end
    end
  end

  def find_hershim(file)
    if file =~ /mlt_o_rav_(\d\d\d\d-\d\d-\d\d)_lesson_full/ || file =~ /mlt_o_rav_.*_(\d\d\d\d-\d\d-\d\d)_lesson/
      date                          = $1
      @event.morning_lesson_hershim = hershim = @hershim_list.grep(/shiur-boker-helek-.+_#{date}/).any?
    else
      hershim_name                  = file.gsub(/_hrsh$|_hrsh-[\da-f]+$|-[\da-f]+$/, '')
      hershim                       = @hershim_list.grep(/#{hershim_name}$/).any?
      @event.morning_lesson_hershim = false
    end
    $stderr.puts "HERSHIM #{hershim ? 'MATCH' : 'MISS '} #{file}" if $verbose
    @event.hershim = hershim
  end

end

opts = GetoptLong.new(
    ['--start', '-s', GetoptLong::REQUIRED_ARGUMENT],
    ['--finish', '-f', GetoptLong::REQUIRED_ARGUMENT],
    ['--playlists', '-p', GetoptLong::OPTIONAL_ARGUMENT],
    ['--hershim', '-e', GetoptLong::OPTIONAL_ARGUMENT],
    ['--help', '-h', GetoptLong::NO_ARGUMENT],
    ['--verbose', '-v', GetoptLong::NO_ARGUMENT],
    ['--super-verbose', '-V', GetoptLong::NO_ARGUMENT],
)

$verbose = false

now    = Time.now
start  = Date.new(now.year)
finish = Date.new(now.year, now.month, now.mday)

playlists_dir = 'playlists'
hershim_dir   = 'hershim'

begin
  opts.each do |opt, arg|
    case opt
      when '--start'
        start = Date.parse(arg)
      when '--finish'
        finish = Date.parse(arg)
      when '--playlists'
        playlists_dir = arg.gsub(/\\/, '/')
      when '--hershim'
        hershim_dir = arg.gsub(/\\/, '/')
      when '--verbose'
        $verbose = true
      when '--super-verbose'
        $super_verbose = true
        $verbose       = true
      else
        exit 1
    end
  end
rescue
  puts <<-EOF
-h, --help:
  show help

-s, --start:
  start date (default: #{start})

-f, --finish:
  finish date (default: #{finish})

-p, --playlist:
  directory with playlists (default: #{playlists_dir})

-e, --hershim:
  directory with files for hershim (default: #{hershim_dir})
  EOF

  exit 1
end

Playlist.new(start, finish, playlists_dir, hershim_dir).parse
exit 0
